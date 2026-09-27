"""Bad-input handling: every failure class must come back as a categorised
exception, never a silent OK and never a silently dropped file."""

import pytest

from excel_consolidation.consolidator import process_directory
from excel_consolidation.generator import build_manifests, generate_all

SEED = 7


@pytest.fixture(scope="session")
def result(tmp_path_factory):
    workbooks = tmp_path_factory.mktemp("wb_fail")
    generate_all(workbooks, seed=SEED)
    out = tmp_path_factory.mktemp("out_fail")
    return process_directory(workbooks, out)


def test_all_30_files_are_accounted_for(result):
    assert len(result.file_results) == 30


def test_no_file_is_silently_dropped(result):
    manifests = build_manifests(SEED)
    assert {fr.filename for fr in result.file_results} == {m.filename for m in manifests}


def test_corrupt_file_is_a_categorised_exception_not_ok(result):
    corrupt = next(fr for fr in result.file_results if fr.filename == "15_corrupt_export.xlsx")
    assert corrupt.status == "unreadable"
    assert corrupt.status != "ok"
    assert corrupt.detail, "the exception type/message must be recorded, not swallowed"
    assert corrupt.rows_in == 0
    assert corrupt.rows_out == 0


def test_wrong_structure_file_is_a_categorised_exception_not_ok(result):
    bad = next(fr for fr in result.file_results if fr.filename == "27_inventory_snapshot.xlsx")
    assert bad.status == "unmappable_structure"
    assert bad.status != "ok"
    assert bad.rows_out == 0
    assert bad.rejected == bad.rows_in
    assert bad.rows_in > 0
    reasons = {rr.reason for rr in result.rejected_rows if rr.file == bad.filename}
    assert reasons == {"unmappable_structure"}


@pytest.mark.parametrize("reason", [
    "missing_order_id",
    "missing_order_date",
    "missing_product",
    "missing_quantity",
    "missing_amount",
    "invalid_date",
    "invalid_quantity",
    "invalid_amount",
    "blank_row",
    "totals_row",
    "unmappable_structure",
])
def test_every_bad_input_class_is_categorised(result, reason):
    reasons = {rr.reason for rr in result.rejected_rows}
    assert reason in reasons, f"no rejected row carried reason={reason!r} in the seed={SEED} corpus"


def test_duplicates_are_reported_with_provenance_not_silently_dropped(result):
    assert len(result.duplicate_rows) > 0
    for dr in result.duplicate_rows:
        assert dr.order_id
        assert dr.first_seen_file
        assert dr.first_seen_row > 0


def test_no_status_is_ok_for_a_file_with_any_rejection_or_duplicate(result):
    for fr in result.file_results:
        if fr.rejected or fr.duplicates:
            assert fr.status != "ok"
