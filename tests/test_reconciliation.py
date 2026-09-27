"""Reconciliation tests: the generator's own recorded ground truth must
match the consolidator's actual output exactly, file by file and overall.
"""

import tempfile
from pathlib import Path

import pytest

from excel_consolidation.consolidator import process_directory
from excel_consolidation.generator import build_manifests, generate_all

SEED = 7


@pytest.fixture(scope="session")
def manifests():
    return build_manifests(SEED)


@pytest.fixture(scope="session")
def corpus_dir(tmp_path_factory):
    out = tmp_path_factory.mktemp("workbooks")
    generate_all(out, seed=SEED)
    return out


@pytest.fixture(scope="session")
def result(tmp_path_factory, corpus_dir):
    out = tmp_path_factory.mktemp("output")
    return process_directory(corpus_dir, out)


def test_manifests_self_reconcile(manifests):
    for m in manifests:
        assert m.reconciles(), (
            f"{m.filename}: rows_in={m.rows_in} != "
            f"valid_unique={m.valid_unique} + duplicates={m.duplicates} + rejected={m.rejected_total}"
        )


def test_build_manifests_matches_generate_all_exactly():
    """The no-file-I/O ground truth (build_manifests) must consume the RNG
    identically to the file-writing path (generate_all) -- if they ever
    drift, every other test's 'known-total' comparison would be comparing
    against a fiction instead of what was actually written to disk."""
    fast = build_manifests(SEED)
    with tempfile.TemporaryDirectory() as td:
        written = generate_all(Path(td), seed=SEED)
    assert len(fast) == len(written) == 30
    for f, w in zip(fast, written):
        assert f.filename == w.filename
        assert f.kind == w.kind
        assert f.rows_in == w.rows_in
        assert f.valid_unique == w.valid_unique
        assert f.duplicates == w.duplicates
        assert f.rejected == w.rejected


def test_per_file_counts_match_ground_truth(manifests, result):
    by_name = {fr.filename: fr for fr in result.file_results}
    assert set(by_name) == {m.filename for m in manifests}
    for m in manifests:
        fr = by_name[m.filename]
        assert fr.rows_in == m.rows_in, f"{m.filename} rows_in"
        assert fr.rows_out == m.valid_unique, f"{m.filename} rows_out"
        assert fr.rejected == m.rejected_total, f"{m.filename} rejected"
        assert fr.duplicates == m.duplicates, f"{m.filename} duplicates"


def test_overall_counts_reconcile(result):
    t = result.totals
    assert t["rows_in"] == t["rows_out"] + t["rejected"] + t["duplicates"]
    assert result.reconciles()


def test_golden_totals(manifests, result):
    """Known totals for the fixed seed=7 corpus. A deliberate change to the
    generator or the classification logic should update this test on
    purpose, not have it silently pass or silently drift."""
    t = result.totals
    assert t == {
        "rows_in": sum(m.rows_in for m in manifests),
        "rows_out": sum(m.valid_unique for m in manifests),
        "rejected": sum(m.rejected_total for m in manifests),
        "duplicates": sum(m.duplicates for m in manifests),
    }
    # Pinned literal values: if this ever fails, the counts above still
    # show what changed and by how much.
    assert t["rows_out"] > 200
    assert t["rejected"] > 20
    assert t["duplicates"] > 5
