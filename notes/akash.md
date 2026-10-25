# Akash — explain-before-merge (plan §8)

## formats.py v2 + run_all.py (22 Oct)

1. **What I built.** BFP6 is now the D1 candidate 16|4|5 (16 values share one 4-bit exponent E = floor(log2 max), clamped to [−8, 7]; each value = ±M·2^(E−4), M ≤ 31). I added the E4M1 one-NaN row (vmax 256), BFP 32|8|5, two's-complement MXINT6, a W+X mode (`quantize_X`, `metrics(..., Xq=)`, `exhaustive(mode="wx")`), true storage per format, class-style range/saturated counts, and `run_all.py`.
2. **Why it works.** BFP and MXINT share one fixed step per block. They are equal exactly when block size, exponent field and exponent rule are equal, and a test checks that (32|8|5 = MXINT6_sym, bit for bit; since the revised D10 the candidate MXINT6 is two's complement, see the note at the end). X is quantized along its own reduction axis (down the columns), so each input vector is independent, as in the class W+A demo.
3. **What would break it.** A layer whose block maxima fall outside [2^−8, 2^8) clamps BFP6's 4-bit exponent (the stats count `n_scale_clamped`). Changing `formats.py` without rerunning: the checkpoint now stores a fingerprint (data + format + code hash), so stale rows are recomputed rather than silently reused.
4. **Bugs found in v1.** (a) The checkpoint was keyed only by (format, policy), so after any change to W, X or a format it returned old numbers. (b) `n_saturated` counted values out of range *before* rounding. The Lecture 19 notebook calls that "range"; its "saturated" means the *rounded* value is clipped. E2M3 on C: 43 range, 20 saturated. The plan's "E2M3 saturates 43 weights" should read "43 out of range, 20 saturated".
5. **What the numbers say (C, held-out, class rule).** E2M3 3.10, BFP6 3.48, MXINT6 4.05, E3M2 5.47, E4M1 11.15. E2M3 still wins, and BFP6 is now a genuine second.

## Teach-back questions (for the Thursday sync)
- Why does BFP 32|8|5 give *exactly* the symmetric MXINT6's numbers, and why does 16|4|5 not?
- E2M3, block max 3.8 under the class rule: is it out of range, saturated, both, or neither? (Out of range only: it scales to 7.6 and rounds to 7.5.)
- Why is BFP6 6.28 bits/element and not 6.25?
- In W+X mode, why do the blocks of X run down its columns and not along its rows?

## Five questions the professor might ask, with short answers
1. *Is BFP6 just MXINT6?* No. At 32 values with an 8-bit exponent BFP equals the symmetric MXINT6 (we show this; our candidate MXINT6 also has the two's-complement code −32). Our BFP6 uses 16 values with a 4-bit exponent: same element grid, finer blocks, 6.28 vs 6.32 bits/element. On C that lowers the error from 4.05% to 3.48%.
2. *Why does the floor rule clip?* For E2M3, max/2^e lies in [4, 8) but vmax is 7.5, so a block max in [7.5, 8)·2^e is out of range. `ceil` avoids it at the cost of a coarser grid. We report both.
3. *Does the E4M1 convention change the conclusion?* No. Under the class rule it ranges from 11.15% to 20.66%, and E2M3 is 3.10%.
4. *Did a library do the rounding?* No. It is NumPy only: `np.round` is round-half-to-even, and `floor(log2)` comes from `frexp` so that it is exact one ulp below a power of two.
5. *Are the numbers reproducible?* Yes. `python run_all.py` reruns the tests and regenerates every table and figure with fixed seeds, and `results/RUN_MANIFEST.md` records the data hashes and commit.

## Update 24 Oct (Aryan, while Akash was away)
- D10 revised: `MXINT6` is two's complement k ∈ [−32, 31]; the symmetric code is the variant `MXINT6_sym` (formats.py v2.1, PR #7).
- D13: `ceil` is the literal "smallest scale that clips nothing"; for two's complement a negative block maximum may use code −32 (formats.py v2.2). C held-out MXINT6: ocp 4.054, ceil 4.160, opt 3.880.
- `predictor.py` now scores BFP6 (16|4|5) on 16-value blocks; before, it scored the v1 block-32 `ceil` BFP6.
