# Final check of code and report (Siddhant), 25 Oct

Commit: `3ad35b3` (merge of PR #11, 25 Oct 07:44). Report: `report/report.docx` (built at `58eba9e`).

## 1. Code

| Check | Result |
|---|---|
| Bit-exact cross-check, `formats.py` v2.2 vs `verification/reference_impl.py` (adversarial + 5 random families + W_C, all formats, variants, policies) | **217 / 217** |
| D13 change to `reference_impl.py` (`ceil` for two's-complement MXINT6, added by Aryan) | reviewed: positives must fit 31/16·2^s, negatives 32/16·2^s, smallest such s. Correct. Reference gives MXINT6 ceil = 4.160% held-out, as in the report |
| Clean environment: `git archive` copy, new venv, `pip install -r requirements.txt`, `python run_all.py --fresh` (Linux, Python 3.13.16, numpy 2.5.4) | **exit 0, 35 s** |
| Tests | `test_formats.py` 31/31, `test_scale_opt.py` 16/16, `test_selector.py` 12/12, `verification/test_reference_impl.py` 16/16 |
| Regenerated vs committed | T1–T6, `block_study.md`, `q3_study.md`, `scale_study.md`, `selector_bench.md`, `selector_scaling.md`, `x_stress.md`: **identical**. `selector_bench.csv` (26,858 rows): identical. `baseline_C*.csv`: counts identical, floats ≤ 6.4·10⁻¹⁶ relative (committed run on macOS / numpy 2.5.1) |
| T7, T8 (datasets A–E) | need `assignment_data.npz`; without it `run_all.py` now marks T7 "not regenerated" (fixed). Not re-verified here |
| Earlier issues (24 Oct) | T4 label fixed (`= MXINT6_sym`), stray json removed, T7 marked when stale, `requirements.txt` added: **all fixed** |

## 2. Report numbers

Every decimal number in `report.docx` (662) was matched against the regenerated `results/` files: **662 / 662 found** (plus 4 arXiv IDs). The only one not printed verbatim, "0.0004 pp" (B: MXINT6 vs BFP6), is 0.000413 in `T7_datasets_A_to_E.csv`. Key counts checked in their source files: flushed 461 / 748 / 891 / 4282 (T6), 160 / 1066 / 1904 / 1289 (`q3_study.md`), 506 of 524,288 blocks and 1.698 s vs 2.046 s (`selector_scaling.md`), 8 of 700 clamped BFP blocks (`q3_study.md`). Recomputed independently: W_C kurtosis 5.76, max 5.70σ.

## 3. Before submitting (text, not numbers)

1. **Roll numbers.** Line 4 of the report is still the placeholder "[Roll numbers: add them in the submitted copy only …]". Add them to the submitted copy.
2. **Storage accounting of the Lecture 19 notebook.** The report gives only 6.32 / 6.28 bits per element, and Table 1 tags "a tail of 4 with its own scale" as [CLASS]. The notebook zero-pads the tail (same values, so [CLASS] is right for the values), but its storage estimate (`ideal_weight_size`) stores every block in full: **8.00** bits/element for MX and **7.00** for BFP6 16|4|5 at K = 100. Suggested sentence under Table 2: *"The Lecture 19 notebook counts the padded tail as stored, which gives 8.00 (MX) and 7.00 (BFP6) bits per element; we report the real storage (D11). The element width is 6 bits in every case, and the ranking does not depend on this choice."*
3. **"Same integer grid" (Sections 3 and 5.3, dataset B).** BFP6 (16 values, sign-magnitude, k ≤ 31) and MXINT6 (32 values, two's complement, k ≥ −32) do not have the same grid; they have the same step relative to the block maximum. Suggest: "which use the same integer step relative to their block maximum". Table 5 also prints "BFP6 / 0.000" for B while the text calls it a tie: write "tie" in the winner cell, as D6 says.
4. **Section 7.** Suggest adding: *"Siddhant's reference was written from the conventions without reading formats.py; the D13 `ceil` rule for two's complement was added to it later and reviewed. Every [CLASS] convention was checked against the Lecture 19 notebook (verification/CONVENTIONS_AND_STORAGE.md)."*
5. **References.** The arXiv IDs and venues are consistent with what I know, but the team rule (plan §8) is that every citation is opened before submission.
6. **AI use.** If the course requires declaring AI tools, the report needs a line for it.
