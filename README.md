# CS611 Assignment 1 — Group 8: dataset C, 6-bit element width

Data-dependent selection of a 6-bit block format (MXFP6 E3M2, MXFP6 E2M3, MXFP6 E4M1, MXINT6, BFP6)
for the weight matrix W_C and its calibration input X_cal.

| Member | Roll number |
|---|---|
| Aryan Kumar | 261110012 |
| Akash Kumar | 261110008 |
| Nikita Kumari Pandit | 261110078 |
| Siddhant Singh Rathoria | 261110062 |
| Anuj Chaudhary | 210167 |

## What is in this package

**Start with the report, `Report_Group8_DatasetC_6bit.docx`, at the root of the zip.** It is
self-contained: formats, assumptions, methods, results and their sources. Everything else is the
code and data that produce its numbers. The report names files by their path inside this zip.

| Path | What it is |
|---|---|
| `Report_Group8_DatasetC_6bit.docx` | the report |
| `README.md` | this file |
| `requirements.txt` | Python packages |
| `data/` | `W_C.npy` (100×100) and `X_cal.npy` (100×256) |
| `formats.py` | the five quantizers and their variants, scale policies, metrics, exhaustive harness |
| `selector.py` | **the chosen selection method** (`select_format`) and every method it is compared with |
| `selector_suite.py` | test matrices for the selectors: the course datasets, 28 adversarial families, 1000 random matrices with 9 activation families |
| `run_all.py`, `run_baseline.py` | one command for the main tables T1–T9 (`results/tables/`) and the baseline runs |
| `bench_selector.py`, `bench_scaling.py`, `x_stress.py` | Q2: accuracy against cost, cost against matrix size, selection under different activations |
| `q3_study.py`, `block_study.py` | Q3: outliers, sparsity, variation in scale; block length at equal storage |
| `scale_opt.py`, `scale_study.py` | the scale-optimization study, with an independent implementation of the scale policies |
| `make_figures.py`, `make_selector_figures.py`, `explore.py`, `predictor.py`, `validate_predictor.py` | figures, data exploration, the histogram predictor (a compared method) |
| `test_formats.py`, `test_selector.py`, `test_scale_opt.py` | unit tests |
| `verification/` | an exact-fraction reference implementation, its tests and a cross-check harness |
| `results/` | every generated table and log (`results/tables/T1..T9`, `*_wx.md` = W and X quantized, `*_w.md` = weights only) |
| `figures/` | every figure in the report |
| `report/build_report.py` | builds the report from `results/` (every number in it is read from there) |

## Answer in one paragraph

**MXFP6 E2M3** is the most suitable 6-bit format for dataset C. With W and X both quantized it has
the lowest held-out output error under every scale rule we tried (4.244% against MXINT6 5.213% and
BFP6 5.218% under the class rule), and it is also first when only the weights are quantized. It
stays first under all 11 activation families we tried.
The chosen selection method, an **exact activation-weighted sketch**, scores all five candidates
without quantizing anything and makes the same choice as brute force with the same objective.
Report sections 3–7 give the details.

## How to reproduce

Python 3.8+ with numpy and matplotlib (`tqdm` optional, `python-docx` only for rebuilding the report):

    python -m pip install -r requirements.txt

The other course datasets (A–E, used only for context) are in the course file `assignment_data.npz`;
pass it with `--npz path/to/assignment_data.npz`. Without it, everything for dataset C still runs.

    python run_all.py --npz assignment_data.npz         # tests, baselines, tables T1-T9, figures (~30 s)
    python bench_selector.py --npz assignment_data.npz  # Q2, W and X quantized (add --mode w for weights only; ~5 min)
    python bench_scaling.py                              # Q2, selection cost against matrix size
    python x_stress.py --npz assignment_data.npz        # selection under 11 activation families
    python q3_study.py                                   # Q3 (about 2 min)
    python block_study.py --npz assignment_data.npz     # block length at equal storage
    python make_selector_figures.py                      # figures 6 and 7
    python report/build_report.py                        # rebuild the report from results/

Tests:

    python test_formats.py                       # 31 tests
    python test_selector.py                      # 13 tests, including the exactness of the chosen method
    python test_scale_opt.py                     # 16 tests
    python verification/test_reference_impl.py   # 17 tests

## Using the chosen method on other matrices

    import selector as S
    best, scores = S.select_format(W, X_cal)            # W and X quantized (default)
    best, scores = S.select_format(W, X_cal, mode="w")  # weights only, X kept in full precision

`scores` is the predicted squared output error of each candidate; `best` is the smallest.

## Secondary reference

The same code is in the group's GitHub repository, https://github.com/akashiitkanpur/cs611-g8
(private; access on request). This zip is the submission.
