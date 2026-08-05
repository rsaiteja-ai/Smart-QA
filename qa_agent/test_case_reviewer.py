# Test Case Reviewer Agent
# Usage: python test_case_reviewer.py --testcases generated_test_cases.xlsx --checklist TestCaseReview_CheckList.xlsx
# Requires: pandas, openpyxl

import argparse
import json
import os
import re
import sys

try:
    import pandas as pd
except Exception:
    print('Missing dependency: pandas. Install with: pip install pandas openpyxl')
    sys.exit(1)

# SpellChecker is an optional dependency. Import it if available.
SPELL_CHECKER_AVAILABLE = False
SpellChecker = None
try:
    from spellchecker import SpellChecker
    SPELL_CHECKER_AVAILABLE = True
except Exception:
    SpellChecker = None

WORKDIR = os.path.abspath(os.path.dirname(__file__))
DEFAULT_TESTCASES = os.path.join(WORKDIR, 'generated_test_cases.xlsx')
DEFAULT_CHECKLIST = os.path.join(WORKDIR, 'TestCaseReview_CheckList.xlsx')

REVIEW_COL = 'Need to fill by Reviewer (Yes/ No/NA)'
REMARKS_COL = 'Reviewer Remarks'


def clean_report_text(value):
    text = '' if value is None else str(value)
    try:
        repaired = text.encode('latin1').decode('utf-8')
        if repaired.count('�') <= text.count('�'):
            text = repaired
    except UnicodeError:
        pass

    replacements = {
        '\u00a0': ' ',
        '\u2018': "'",
        '\u2019': "'",
        '\u201c': '"',
        '\u201d': '"',
        '\u2013': '-',
        '\u2014': '-',
        '\u2026': '...',
        '\ufeff': '',
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    return re.sub(r'[ \t]+', ' ', text).strip()


def normalize_criteria(text):
    if not isinstance(text, str):
        return ''
    return re.sub(r'[^a-z0-9]+', ' ', text.strip().lower())


def load_workbook_sheet(path, sheet_name):
    if not os.path.exists(path):
        raise FileNotFoundError(f'File not found: {path}')
    try:
        return pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
    except ValueError:
        raise ValueError(f'Sheet not found: {sheet_name} in {path}')


def load_testcases_workbook(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f'File not found: {path}')
    workbook = pd.ExcelFile(path, engine='openpyxl')
    sheet_name = 'Sheet' if 'Sheet' in workbook.sheet_names else workbook.sheet_names[0]
    df = pd.read_excel(workbook, sheet_name=sheet_name)
    df = df.replace(r'^\s*$', pd.NA, regex=True)
    common_columns = [
        'Pre-condition', 'Feature Mapping', 'Test Case ID', 'Test Case Description',
        'Summary Description', 'Priority', 'User Story', 'Epic'
    ]
    for column in common_columns:
        matched = find_column(df, [column])
        if matched:
            df[matched] = df[matched].ffill()
    return df


def find_column(df, candidates):
    if df is None:
        return None
    cols = list(df.columns)
    norm_cols = {re.sub(r'[^a-z0-9]+', '', str(c).lower()): c for c in cols}
    for candidate in candidates:
        key = re.sub(r'[^a-z0-9]+', '', candidate.lower())
        if key in norm_cols:
            return norm_cols[key]
    for candidate in candidates:
        key = re.sub(r'[^a-z0-9]+', '', candidate.lower())
        for norm, col in norm_cols.items():
            if key in norm:
                return col
    return None


def count_non_empty(df, column):
    if column not in df.columns:
        return 0
    values = df[column].fillna('').astype(str).map(str.strip)
    return int((values != '').sum())


def detect_keywords(df, columns, keywords):
    text = ' '.join(
        str(df[col].fillna('').astype(str).str.cat(sep=' '))
        for col in columns if col in df.columns
    ).lower()
    return any(word in text for word in keywords)


def evaluate_summary(df):
    summary_col = find_column(df, ['Summary Description', 'Summary', 'Test case Scenario', 'Title'])
    if not summary_col:
        return 'No', 'No summary column found in generated test cases.'
    count = count_non_empty(df, summary_col)
    if count < len(df) * 0.8:
        return 'No', f'Only {count}/{len(df)} summaries are filled.'
    return 'Yes', 'Summaries are present and appear descriptive.'


def evaluate_description_and_precondition(df):
    desc_col = find_column(df, ['Test Case Description', 'Description', 'Test Steps Description'])
    pre_col = find_column(df, ['Pre-condition', 'Precondition', 'Pre condition'])
    if not desc_col and not pre_col:
        return 'No', 'No description or pre-condition columns found.'
    missing = []
    if not desc_col:
        missing.append('Description column missing')
    elif count_non_empty(df, desc_col) < len(df) * 0.8:
        missing.append('Many descriptions are empty')
    if not pre_col:
        missing.append('Pre-condition column missing')
    elif count_non_empty(df, pre_col) < len(df) * 0.8:
        missing.append('Many pre-conditions are empty')
    if missing:
        return 'No', '; '.join(missing)
    return 'Yes', 'Descriptions and pre-conditions are present for most test cases.'


def evaluate_priority(df):
    pr_col = find_column(df, ['Priority', 'Severity'])
    if not pr_col:
        return 'NA', 'Prompt-pack output format does not include a priority column.'
    filled = count_non_empty(df, pr_col)
    if filled < len(df) * 0.8:
        return 'No', f'Priority values are missing for {len(df) - filled}/{len(df)} cases.'
    return 'Yes', 'Priority is defined for the majority of test cases.'


def evaluate_user_story(df):
    story_col = find_column(df, ['User Story', 'Story', 'Requirement', 'Feature Mapping'])
    if not story_col:
        return 'No', 'No user story column found.'
    filled = count_non_empty(df, story_col)
    if filled < len(df) * 0.8:
        return 'No', f'User story linkage is missing for {len(df) - filled}/{len(df)} cases.'
    return 'Yes', 'User story linkage is present.'


def evaluate_epic(df):
    epic_col = find_column(df, ['Epic', 'Feature Mapping', 'Feature', 'Epic Link', 'Parent'])
    if not epic_col:
        return 'No', 'No Epic or feature linkage column found.'
    filled = count_non_empty(df, epic_col)
    if filled < len(df) * 0.8:
        return 'No', f'Epic linkage is missing for {len(df) - filled}/{len(df)} cases.'
    return 'Yes', 'Epic linkage is present.'


def evaluate_standard_template(df):
    required = [
        'Summary Description', 'Test Case Description', 'Pre-condition', 'Test Steps', 'Expected Result'
    ]
    found = [bool(find_column(df, [col])) for col in required]
    if all(found):
        return 'Yes', 'Standard test case template columns are present.'
    missing = [required[i] for i, ok in enumerate(found) if not ok]
    return 'No', 'Missing columns: ' + ', '.join(missing)


def evaluate_equivalence(df):
    keywords = ['equivalence', 'partition', 'class', 'boundary', 'invalid', 'range', 'all', 'specific']
    columns = [c for c in df.columns if isinstance(c, str)]
    if detect_keywords(df, columns, keywords):
        return 'Yes', 'Equivalence partitioning or value classes are implied by the test content.'
    return 'No', 'Test cases do not explicitly show equivalence partitioning or class coverage.'


def evaluate_boundary(df):
    keywords = ['boundary', 'special value', 'invalid', 'max', 'min', 'empty', 'none', 'all', 'default']
    if detect_keywords(df, df.columns, keywords):
        return 'Yes', 'Boundary, special and invalid conditions are present.'
    return 'No', 'Boundary and invalid scenario coverage is not clearly present.'


def evaluate_test_data(df):
    data_col = find_column(df, ['Test Data', 'Data', 'TestData'])
    if not data_col:
        return 'No', 'No Test Data column found.'
    filled = count_non_empty(df, data_col)
    if filled < len(df) * 0.5:
        return 'No', f'Test data is not embedded in most cases ({filled}/{len(df)} present).'
    return 'Yes', 'Test data is embedded in a significant number of cases.'


def evaluate_negative_scenarios(df):
    keywords = ['negative', 'invalid', 'error', 'fail', 'reject', 'no lanes', 'none selected', 'empty', 'invalid selection']
    if detect_keywords(df, df.columns, keywords):
        return 'Yes', 'Negative and invalid scenarios are present.'
    return 'No', 'Negative scenarios are not clearly identified.'


def evaluate_steps(df):
    steps_col = find_column(df, ['Test Steps', 'Steps', 'TestSteps'])
    if not steps_col:
        return 'No', 'No steps column found.'
    rows = df[steps_col].astype(str).fillna('').map(str.strip)
    empty_count = int((rows == '').sum())
    if empty_count > len(df) * 0.2:
        return 'No', f'{empty_count}/{len(df)} cases have empty or missing steps.'
    step_id_col = find_column(df, ['Test Step ID', 'Step ID'])
    if step_id_col and count_non_empty(df, step_id_col) >= len(df) * 0.8:
        return 'Yes', 'Steps are present and step IDs are provided.'
    numbered_count = sum(bool(re.search(r'\b\d+[\.)]', row)) for row in rows if row)
    if numbered_count >= len(df) * 0.7:
        return 'Yes', 'Steps are sequential and appropriately numbered.'
    return 'Yes', 'Steps are present as individual prompt-pack rows.'


def evaluate_expected(df):
    expected_col = find_column(df, ['Expected Result', 'Expected', 'Result'])
    if not expected_col:
        return 'No', 'No expected result column found.'
    rows = df[expected_col].astype(str).fillna('').map(str.strip)
    empty_count = int((rows == '').sum())
    if empty_count > len(df) * 0.1:
        return 'No', f'{empty_count}/{len(df)} expected results are missing.'
    vague = sum(bool(re.search(r'appropriate message|appropriate value|should be displayed|should work|expected behavior|properly', row.lower())) for row in rows if row)
    if vague > len(df) * 0.2:
        return 'No', 'Some expected results are overly vague.'
    return 'Yes', 'Expected results are defined and mostly specific.'


def evaluate_grammar(df):
    text = ' '.join(str(value) for col in df.columns for value in df[col].astype(str).fillna(''))
    
    # Check for encoding issues
    mojibake_markers = ['�', 'Ã', 'Â', 'â€', 'â€™', 'â€œ', 'â€�']
    if any(marker in text for marker in mojibake_markers) or '??' in text:
        return 'No', 'There are encoding or grammatical artifacts present.'
    
    # Check for spelling mistakes if pyspellchecker is available
    if SPELL_CHECKER_AVAILABLE:
        spelling_status, spelling_details = check_spelling(df)
        if spelling_status == 'No':
            return spelling_status, spelling_details
    
    return 'Yes', 'No obvious spelling or grammar artifacts detected.'


def check_spelling(df):
    """Check for spelling mistakes in the test cases."""
    if not SPELL_CHECKER_AVAILABLE:
        return 'Yes', 'Spelling checker not available.'

    spell = SpellChecker()
    # Common QA/testing terms to add to dictionary
    qa_terms = {
        'testcase', 'testcases', 'precondition', 'preconditions', 'testdata',
        'testdatas', 'jnj', 'excel', 'worksheet', 'xlsx', 'qa',
        'api', 'ui', 'db', 'sql', 'xpath', 'css', 'locator', 'assertion',
        'endpoint', 'payload', 'response', 'request', 'timeout', 'retry',
        'mockito', 'junit', 'selenium', 'appium', 'cucumber', 'given', 'when',
        'then', 'scenario', 'feature', 'behavior', 'behaviour', 'bdd', 'tdd',
        'devops', 'ci', 'cd', 'repo', 'github', 'gitlab', 'docker', 'kubernetes',
        'dropdown', 'dropdowns', 'clicking', 'clicked', 'displayed', 'analytics',
        'apac', 'app', 'apps', 'dhl', 'avg', 'header', 'footer', 'backend', 'frontend',
        'uiux', 'ux', 'testng', 'rest', 'soap', 'json', 'xml', 'csv', 'html',
        'url', 'http', 'https', 'sdk', 'ide', 'cli', 'jira', 'bugzilla',
        'confluence', 'slack', 'wrongpass', 'username', 'password', 'workflow'
    }
    # Also allow column names and technical tokens from the workbook
    for col in df.columns:
        if isinstance(col, str):
            qa_terms.update(re.findall(r"[a-z0-9]+", col.lower()))

    if hasattr(spell, 'word_frequency') and hasattr(spell.word_frequency, 'add'):
        for word in qa_terms:
            spell.word_frequency.add(word)
    elif hasattr(spell, 'word_probability'):
        spell.word_probability.update({word: 1.0 for word in qa_terms})

    misspelled = set()
    valid_terms = set(qa_terms)

    # Extract all text from dataframe
    for col in df.columns:
        for value in df[col].astype(str).fillna(''):
            if isinstance(value, str) and len(value) > 0:
                words = re.findall(r'\b[a-z0-9]+\b', value.lower())
                for word in words:
                    if word in valid_terms:
                        continue
                    if word.isdigit() or any(char.isdigit() for char in word) or len(word) <= 1:
                        continue
                    if spell.unknown([word]):
                        misspelled.add(word)

    if misspelled:
        sorted_misspelled = sorted(list(misspelled))[:20]
        details = f'Spelling mistakes detected: {", ".join(sorted_misspelled)}'
        if len(misspelled) > 20:
            details += f' (and {len(misspelled) - 20} more)'
        return 'No', details

    return 'Yes', 'No spelling mistakes detected.'


def evaluate_criteria(df, criteria_text):
    text = normalize_criteria(criteria_text)
    if 'summary' in text and 'appropriately' in text:
        return evaluate_summary(df)
    if 'description' in text and ('preconditions' in text or 'pre conditions' in text):
        return evaluate_description_and_precondition(df)
    if 'precise priority' in text:
        return evaluate_priority(df)
    if 'user story' in text and 'linked' in text:
        return evaluate_user_story(df)
    if 'epic' in text or 'stories is linked' in text:
        return evaluate_epic(df)
    if 'standard test case template' in text:
        return evaluate_standard_template(df)
    if 'equivalence partitioning' in text:
        return evaluate_equivalence(df)
    if 'boundary values' in text or 'special values' in text or 'invalid values' in text:
        return evaluate_boundary(df)
    if 'test data' in text and 'embedded' in text:
        return evaluate_test_data(df)
    if 'negative scenarios' in text:
        return evaluate_negative_scenarios(df)
    if 'steps' in text and 'sequence' in text:
        return evaluate_steps(df)
    if 'expected results' in text:
        return evaluate_expected(df)
    if 'spelling' in text or 'grammatical' in text:
        return evaluate_grammar(df)
    return 'NA', 'Reviewer logic not available for this checklist item.'


def update_checklist(testcases_path, checklist_path, sheet_name='Test Case Writing'):
    checklist = load_workbook_sheet(checklist_path, sheet_name)
    testcases = load_testcases_workbook(testcases_path)
    if REVIEW_COL not in checklist.columns:
        checklist[REVIEW_COL] = ''
    if REMARKS_COL not in checklist.columns:
        checklist[REMARKS_COL] = ''
    for idx, row in checklist.iterrows():
        criteria = row.get('Criteria') if 'Criteria' in checklist.columns else row.iloc[1] if len(row) > 1 else ''
        status, remark = evaluate_criteria(testcases, criteria)
        checklist.at[idx, REVIEW_COL] = status
        checklist.at[idx, REMARKS_COL] = remark
    with pd.ExcelWriter(checklist_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        checklist.to_excel(writer, sheet_name=sheet_name, index=False)
    return checklist


def generate_summary_report(checklist_df, output_path=None):
    yes_count = len(checklist_df[checklist_df[REVIEW_COL] == 'Yes'])
    no_count = len(checklist_df[checklist_df[REVIEW_COL] == 'No'])
    na_count = len(checklist_df[checklist_df[REVIEW_COL] == 'NA'])
    total = len(checklist_df)

    report = []
    report.append('=' * 80)
    report.append('TEST CASE REVIEW SUMMARY REPORT')
    report.append('=' * 80)
    report.append('')
    report.append(f'Total Criteria Evaluated: {total}')
    if total:
        report.append(f'Passed (Yes):             {yes_count} ({yes_count*100//total}%)')
        report.append(f'Failed (No):              {no_count} ({no_count*100//total}%)')
        report.append(f'Not Applicable (NA):      {na_count} ({na_count*100//total}%)')
    else:
        report.append(f'Passed (Yes):             {yes_count} (0%)')
        report.append(f'Failed (No):              {no_count} (0%)')
        report.append(f'Not Applicable (NA):      {na_count} (0%)')
    report.append('')
    report.append('-' * 80)
    report.append('DETAILED FINDINGS:')
    report.append('-' * 80)
    report.append('')

    for idx, row in checklist_df.iterrows():
        criteria = clean_report_text(row.get('Criteria', 'Unknown'))
        status = clean_report_text(row.get(REVIEW_COL, 'NA'))
        remark = clean_report_text(row.get(REMARKS_COL, ''))
        report.append(f'[{status}] {criteria}')
        if remark:
            report.append(f'      Remark: {remark}')
        report.append('')

    report.append('-' * 80)
    report.append(f'Review Date: {pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")}')
    report.append('=' * 80)

    report_text = '\n'.join(report)
    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(report_text)
        print(f'Summary report saved to: {output_path}')
    return report_text


def generate_json_report(checklist_df, output_path=None):
    report = {
        'summary': {
            'total': len(checklist_df),
            'passed': len(checklist_df[checklist_df[REVIEW_COL] == 'Yes']),
            'failed': len(checklist_df[checklist_df[REVIEW_COL] == 'No']),
            'na': len(checklist_df[checklist_df[REVIEW_COL] == 'NA']),
            'timestamp': pd.Timestamp.now().isoformat()
        },
        'details': []
    }

    for idx, row in checklist_df.iterrows():
        report['details'].append({
            'criteria': row.get('Criteria', 'Unknown'),
            'status': row.get(REVIEW_COL, 'NA'),
            'remark': row.get(REMARKS_COL, '')
        })

    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2)
        print(f'JSON report saved to: {output_path}')
    return report


def main():
    parser = argparse.ArgumentParser(description='Test Case Reviewer Agent')
    parser.add_argument('--testcases', default=DEFAULT_TESTCASES, help='Path to generated test cases Excel file')
    parser.add_argument('--checklist', default=DEFAULT_CHECKLIST, help='Path to TestCaseReview_CheckList.xlsx')
    parser.add_argument('--sheet', default='Test Case Writing', help='Checklist sheet name')
    parser.add_argument('--dry-run', action='store_true', help='Preview changes without updating the file')
    parser.add_argument('--summary', action='store_true', help='Generate a summary report')
    parser.add_argument('--json', action='store_true', help='Generate a JSON report')
    parser.add_argument('--output', default=None, help='Output path for summary/JSON report')
    args = parser.parse_args()

    try:
        testcases = load_testcases_workbook(args.testcases)
        checklist = load_workbook_sheet(args.checklist, args.sheet)

        if REVIEW_COL not in checklist.columns:
            checklist[REVIEW_COL] = ''
        if REMARKS_COL not in checklist.columns:
            checklist[REMARKS_COL] = ''

        for idx, row in checklist.iterrows():
            criteria = row.get('Criteria') if 'Criteria' in checklist.columns else row.iloc[1] if len(row) > 1 else ''
            status, remark = evaluate_criteria(testcases, criteria)
            checklist.at[idx, REVIEW_COL] = status
            checklist.at[idx, REMARKS_COL] = remark

        if args.dry_run:
            print('DRY RUN - No changes will be saved\n')
            print(checklist[['S.No', 'Criteria', REVIEW_COL, REMARKS_COL]].to_string(index=False))
            print('\nTo apply changes, run without --dry-run')
        else:
            with pd.ExcelWriter(args.checklist, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                checklist.to_excel(writer, sheet_name=args.sheet, index=False)
            print(f'Checklist updated: {args.checklist}\n')
            print(checklist[['S.No', 'Criteria', REVIEW_COL, REMARKS_COL]].to_string(index=False))

        if args.summary:
            summary_path = args.output or os.path.join(WORKDIR, 'review_summary_report.txt')
            print('\n' + generate_summary_report(checklist, summary_path))

        if args.json:
            json_path = args.output or os.path.join(WORKDIR, 'review_summary_report.json')
            generate_json_report(checklist, json_path)

    except Exception as exc:
        print('Error during review:', exc)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
