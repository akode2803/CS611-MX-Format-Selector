# Format selector (Q2): our method, and the optional additions (Aryan, 25 Oct)

> **Our method, the one we submit for Q2: the exact sketch, H-weighted** (`sketch_h`).
>
>     import selector as S
>     best, scores = S.select_format(W, X_cal)     # scores = predicted error of every candidate
>
> The assignment asks for one selection method; this is it (§1). Everything in §2 is an **optional addition**, not part of the method:
> - brute-force references it is measured against
> - ablations of it
> - extensions
> - the alternatives we tried
>
> They are kept for the accuracy-vs-cost figure (fig6) and to show why we chose this method. None of them is needed to use it.

Code:
- `selector.py`: the method is `select_format`; `METHOD`, and `ROLE` labels every other approach.
- `selector_suite.py`: the permanent test matrices.
- `test_selector.py`: 12 tests.
- `bench_selector.py`: accuracy vs cost.
- `bench_scaling.py`: cost vs matrix size.
- `x_stress.py`: different activations.
- `make_selector_figures.py`: fig6 and fig7.

Everything uses `formats.py` v2.2 (five candidates, MXINT6 two's complement, BFP6 16|4|5).

```
python test_selector.py
python bench_selector.py --npz assignment_data.npz    # about 5 min; results/selector_bench.{csv,md}
python bench_scaling.py                               # results/selector_scaling.md
python x_stress.py --npz assignment_data.npz          # about 1 min; results/x_stress.md
python make_selector_figures.py
```

## 1. Our method: exact sketch, H-weighted

### 1.1 What it optimizes (way 2, the TA's second "site of optimization")
The output error of a quantized matrix is ‖(W − Wq)X‖² = Σᵢ Δᵢᵀ H Δᵢ, with H = XXᵀ. If the rounding errors are independent and zero-mean, its expectation is Σᵢⱼ H_jj Δᵢⱼ². So each weight counts in proportion to the energy of its input channel, h_j = H_jj = Σ_k X_jk².

Choosing by the weight error alone ("way 1") ignores X. It picks the wrong format in 18% of our random matrices, and in 42% of them when a few activation channels are very large (§1.4).

The same reasoning is why AWQ (Lin et al. 2023) takes weight salience from the activations.

### 1.2 The sketch and the kernel
For a block with reference exponent R (= floor(log2 max|w|); for BFP, clamped to its 4-bit field), write each non-zero weight as |w| = μ·2ᵗ with μ ∈ [1, 2), and let d = R − t be its octave offset below the block peak.

Under the class rule every candidate's scale is 2^(R − emax_f). So the quantized magnitude is c_f(d, μ, sign)·2ᵗ, and c_f depends on the format alone. The squared error of one weight is

  4ᵗ(μ − c)² = m2 − 2c·m1 + c²·m0,  with m0 = 4ᵗ, m1 = 4ᵗμ, m2 = w².

**The algorithm:**
1. h_j = Σ_k X_cal[j, k]² for every input channel.
2. One pass over W, for both block layouts (MX blocks of 32, BFP blocks of 16). Bin every weight by (sign, d, 64 mantissa bins of μ) and add h_j·m2, h_j·m1 and h_j·m0 to its bin. This is the **sketch**: 2 × 49 × 64 × 3 numbers per layout, the same for every format.
3. For each candidate: error_f = ⟨kernel_f, sketch⟩. The kernel is a fixed table computed from the format's definition, with no data.
4. Return the argmin.

### 1.3 It is exact
For 6-bit formats every rounding midpoint and saturation point lies on a multiple of 1/32 of a binade. With 64 mantissa bins, c_f is constant on every bin, so the moments give the error exactly.

`test_selector.py` checks this to 1e-9 for all five candidates, weighted and unweighted, on 9 corner-case matrices:
- ties at exact grid values
- the −32 code of MXINT6
- both ends of BFP's exponent clamp
- zero blocks and the tail block of 4
- 16 octaves of dynamic range
- FP32 data

`test_one_method_and_every_approach_has_a_role` checks that `select_format` returns the same scores as quantizing W with all five formats and measuring the diagonal-H error. So the method makes the **same picks as brute force with the diagonal of H**, without quantizing W.

**What follows from exactness:**
- One pass, shared by all candidates. Scoring a format costs O(#bins), whatever the size of W. Adding a format, or another bit width, means writing its kernel, not touching the data.
- The sketch is additive. Sketches of shards (rows, layers, devices) add up, so it streams and merges.
- The error of every format is decomposed by octave offset and mantissa bin, so it explains why a format wins (flushed energy, clipping, step size). Q3 uses this.

### 1.4 Results of our method
Truth is the held-out Y error (odd X columns) under the class rule; H comes from the even columns only (D6). Numbers are from `results/selector_bench.md` and `results/x_stress.md`.

**Datasets A–E (the course data):**
- Our method picks the held-out winner for A, C, D and E.
- On B it picks MXINT6 where BFP6 wins by 0.002 dB (BFP6 1.781% vs MXINT6 1.782%, `results/tables/T7`): a tie between two formats with the same integer grid.
- For C it picks **MXFP6 E2M3** (held-out 3.101%, ahead of BFP6 3.484%).

**Random suite: 1000 matrices of 128×128.**
- Each matrix mixes heavy tails, sparsity, random, channel and block outliers with continuous magnitudes, and row and column gains.
- Each matrix has one of 9 activation families: white, channel gains, outlier channels, low rank, Student-t3, AR(1) ρ = 0.9, ReLU, massive activations (1% of channels ×100) and 70% zeros.

| | Top-1 | Mean regret | p95 | Max | Median time |
|---|---|---|---|---|---|
| **Our method (exact sketch, H-weighted)** | **93.3%** | **0.031 dB** | **0.06 dB** | 3.8 dB | **1.04 ms** |
| brute force, diagonal H (reference) | 93.3% (same picks) | 0.031 dB | 0.06 dB | 3.8 dB | 2.59 ms |
| brute force, weights only (reference, way 1) | 82.3% | 0.475 dB | 3.3 dB | 20.8 dB | 2.54 ms |

**By activation family** (top-1 / mean regret):

| | white | channel gains | outlier channels | low rank | t3 | AR(1) | ReLU | massive | 70% zeros |
|---|---|---|---|---|---|---|---|---|---|
| **our method** | 98.1% / 0.004 | 95.7% / 0.006 | 95.5% / 0.020 | 83.7% / 0.131 | 94.1% / 0.021 | 90.7% / 0.062 | 87.5% / 0.035 | 96.9% / 0.001 | 96.5% / 0.008 |
| weights only | 98.1% / 0.004 | 75.0% / 0.499 | 63.1% / 1.054 | 80.8% / 0.192 | 94.1% / 0.006 | 90.7% / 0.089 | 88.4% / 0.033 | 58.3% / 2.085 | 97.4% / 0.007 |

**Different activations for the course data** (`x_stress.py`; A–E × the real X_cal, bootstrap resamples and the 9 synthetic families, 5 seeds):
- **C:** E2M3 wins under every family (margins 0.24–0.41 pp). Our method finds it in 100% of seeds, except outlier channels (80%).
- **A and E change winner only under massive activations** (A: E2M3 → BFP6; E: BFP6 → E2M3). Our method follows both; weights only finds them in 20% and 40% of seeds.
- **The X_cal-based choice for C costs at most 0.011 dB** on any X_test family. A, D and E can lose up to 0.7, 0.9 and 1.5 dB when X_test is far from X_cal (low rank or massive activations). No selector can avoid that without seeing X_test.

**Cost at scale** (`results/selector_scaling.md`, LLM-like 4096 × 4096 = 16.8M weights):
- In NumPy the sketch takes 2.2 s, about the same as brute force (2.0 s). Both are memory-bound.
- The sketch's advantage is the shared single pass, not raw speed in NumPy.
- The certified extension (§2) needs 0.17 s and rounds 0.10% of the blocks.

### 1.5 Limitations of our method
- **Diagonal H only.** When X is strongly correlated (low rank), the realized error depends on the projection of ΔW onto a few directions, and no expectation predicts it. Top-1 is 83.7% there, against 98.1% for brute force with the full H. The optional `certified_rows_y` (§2) handles this case (97.1%) at 17× the cost.
- **Class rule.** The method scores the class rule `ocp`. The `opt` policy is covered by the optional `sketch_opt_h` (§2).
- **BFP's exponent clamp** is modelled for d ≥ −8. Blocks more than 8 octaves above the clamp are not modelled.
- **Test data.** The random families are ours. Only A–E are course data, and those are synthetic too.

## 2. Optional additions (not part of the method)
Every approach shares the interface `approach(W, h=None) -> {format: predicted error}`. `ROLE` in `selector.py` labels each one, and fig6 shows them all (our method is the star).

Results on the random suite below; datasets and adversarial in `results/selector_bench.md`.

| Role | Approach | Top-1 | Mean regret | ms | What it shows |
|---|---|---|---|---|---|
| **our method** | **`sketch_h`** | **93.3%** | **0.031 dB** | **1.04** | |
| reference | `exhaustive_h` | 93.3% | 0.031 dB | 2.59 | brute force, diagonal H: identical picks, so the method is exact |
| reference | `exhaustive_w` | 82.3% | 0.475 dB | 2.54 | brute force, weight error (way 1) |
| reference | `exhaustive_y` | 96.2% | 0.010 dB | 26.5 | brute force, full Y on the fit columns: the ceiling for anything that sees only X_cal |
| reference | `exhaustive_opt` | 81.2% | 0.503 dB | 10.2 | ANT/MoFQ style: trial quantization with a per-block scale search, weight MSE |
| ablation | `sketch` | 82.3% | 0.475 dB | 1.05 | the method without H (way 1): same picks as weight brute force |
| ablation | `sketch_nb16_h` | 91.8% | 0.047 dB | 0.96 | 16 mantissa bins instead of 64 |
| ablation | `sketch_nb4_h` | 87.1% | 0.084 dB | 0.97 | 4 bins |
| ablation | `sketch_nb1_h` | 80.8% | 0.201 dB | 1.02 | 1 bin per octave |
| ablation | `closed_form(_h)` | 40.9% (42.2%) | 2.3 dB | 0.63 | coarsest level, two statistics ρ and φ: no flush or clip terms |
| extension | `certified_is_h` | 92.5% | 0.043 dB | 3.03 | importance sample of blocks ∝ Σh·4^F, Hansen–Hurwitz estimate, doubled until the winner leads by 3 s.e.; reads a median 13% of blocks (0.10% at 4096²) |
| extension | `certified_h` | 92.7% | 0.040 dB | 12.8 | the same with uniform sampling |
| extension | `sketch_s10_h` | 71.2% | 0.759 dB | 0.37 | the sketch on a fixed 10% of blocks, uncertified |
| extension | `certified_rows_y` | 95.6% | 0.017 dB | 17.3 | full-H brute force on an importance sample of rows ∝ ‖WᵢX‖²; reads a median 41% of rows; fixes low-rank X (97.1%) |
| extension | `sketch_opt_h` | 89.5% (92.8% vs `opt` truth) | 0.036 dB vs `opt` | 26.7 | exact prediction of the `opt` policy: the per-block sketch scores each scale offset |
| alternative | `hybrid_h` | 90.6% | 0.054 dB | 4.79 | M1 with the top octave rounded exactly |
| alternative | `akash_hist_h` | 89.3% | 0.062 dB | 3.91 | Akash's `predictor.py`: expected error per histogram bin, block maxima exact |
| alternative | `m1_h` (`m1`) | 88.4% (80.6%) | 0.078 dB | 2.44 | per-element expected error: step²/12, flush z², clip (z − vmax)² |
| alternative | `learned_tree` | 73.4% | 0.680 dB | 1.24 | CART (depth 4) on sparsity, kurtosis, ρ32, ρ16, φ, flush fraction; trained on 600 separate matrices |
| alternative | `exh_rows10` | 70.1% | 0.838 dB | 1.22 | brute force on 10% of rows |
| alternative | `fit_gauss` | 61.8% | 1.136 dB | 3.04 | fit a Gaussian, score synthetic draws exactly |
| alternative | `fit_t`, `fit_laplace` | 53.7% | 1.781 dB | 4.09 | Student-t (SF4-style) or Laplace fit, same scoring |
| alternative | `ant_rule` | 5.9% | 7.25 dB | 0.29 | shape rule in the spirit of ANT's motivation (kurtosis → int / flint / float); not ANT's algorithm, which is trial quantization |

**What the additions show:**
- **Why way 2:** the way-1 rows are 11 points lower, and fall to 58–63% on outlier channels and massive activations.
- **Why 64 bins:** coarser sketches lose accuracy steadily, down to the closed form at about 41%.
- **Why not a cheaper model:** every expected-error model (M1, hybrid, `predictor.py`) and every distribution fit is less accurate and slower than the exact sketch.
- **Ways to extend the method if needed:**
  - Large matrices: `certified_is_h`, the same pick from 0.10% of the blocks at 16.8M weights.
  - Correlated X: `certified_rows_y`.
  - The `opt` policy: `sketch_opt_h`.
- **Winners in the random suite:** BFP6 682, E2M3 229, E3M2 59, MXINT6 30. The class-rule winner equals the `opt` winner in 93.7% of matrices.

**Per-block mixed formats** (excluded by D4; `results/block_study.md`). Letting each 32-block pick its own MX format, at +2 bits per block (6.40 bits/element), lowers the held-out error by 0.3–3.6% (C: 3.101% → 2.992%). MixFP4 (Zou et al. 2026) picks E2M1 or E1M2 per block and stores the choice in the unused sign bit of the FP8 E4M3 scale. Our E8M0 scale has no free bit.

## 3. Ideas from the papers
- **AWQ** (Lin et al. 2023): salience comes from activations, not weights. This supports way 2.
- **ANT** (Guo et al., MICRO 2022) and **MoFQ** (Zhang et al. 2023) choose the type per tensor or layer by trial quantization and minimum MSE; ANT adds a clipping-range search. This is our `exhaustive_opt` reference; our method gives the class-rule part of it without quantizing.
- **Linearity theorem / HIGGS** (Malinovskii et al. 2024): per-layer ℓ2 error relates directly to the perplexity increase, and per-layer levels under a size budget are solved by dynamic programming. With exact per-format errors per matrix, choosing one format per matrix across a model under a bit budget becomes the same dynamic program, still inside D4.
- **Kuzmin et al. 2022 (FP8):** the best exponent/mantissa split is "driven by the severity of outliers". Our kernel is the data-exact version of their analytic MSE.
- **van Baalen et al. 2023 (FP8 vs INT8):** a theoretical INT vs FP comparison. **SF4** (Dotzel et al. 2024): Student-t fits for weights, our `fit_t`.

## 4. Timing note
NumPy on one laptop. The 128×128 timings include Python overhead, which dominates for the sampled methods at that size.
