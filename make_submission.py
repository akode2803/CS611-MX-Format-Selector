"""Build the submission zip:  python make_submission.py  ->  dist/CS611_A1_Group8_DatasetC_6bit.zip

The report, the notebook and README.md sit at the root of the zip, next to the code, data, results and
figures that produce the report's numbers.  Working notes, team documents and checkpoints are
left out.  Run run_all.py, the studies and report/build_report.py first.
"""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPORT = "Report_Group8_DatasetC_6bit.docx"
OUT = os.path.join(HERE, "dist", "CS611_A1_Group8_DatasetC_6bit.zip")

ROOT_FILES = [REPORT, "Notebook_Group8_DatasetC_6bit.ipynb", "README.md", "requirements.txt", "FORMATS_API.md",
              "formats.py", "selector.py", "selector_suite.py", "predictor.py", "data.py", "scipy_free_stats.py",
              "run_all.py", "run_baseline.py", "explore.py", "make_figures.py", "make_selector_figures.py",
              "validate_predictor.py", "bench_selector.py", "bench_scaling.py", "x_stress.py", "q3_study.py",
              "block_study.py", "scale_opt.py", "scale_study.py",
              "test_formats.py", "test_selector.py", "test_scale_opt.py", "make_submission.py"]
DIRS = {"data": (".npy",), "results": (".md", ".csv", ".json", ".txt"), "figures": (".png",)}
EXTRA = ["verification/reference_impl.py", "verification/test_reference_impl.py", "verification/crosscheck.py",
         "verification/formats_adapter.py", "verification/verify_C.py", "report/build_report.py"]
SKIP = ("run_all_state.json", ".ckpt.jsonl")


def files():
    out = list(ROOT_FILES) + list(EXTRA)
    for d, exts in DIRS.items():
        for root, _, fs in os.walk(os.path.join(HERE, d)):
            for f in sorted(fs):
                if f.endswith(exts) and not f.endswith(SKIP):
                    out.append(os.path.relpath(os.path.join(root, f), HERE))
    return out


def main():
    paths = files()
    missing = [p for p in paths if not os.path.exists(os.path.join(HERE, p))]
    if missing:
        raise SystemExit("missing: " + ", ".join(missing))
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            z.write(os.path.join(HERE, p), p)
    print(f"wrote {os.path.relpath(OUT, HERE)}: {len(paths)} files, {os.path.getsize(OUT) / 1e6:.1f} MB")


if __name__ == "__main__":
    main()
