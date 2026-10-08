# buckingham-pi-study-notes
Study notes on Buckingham Pi Theorem in PowerPoint format

## Contents

- `generate_presentation.py` — builds the Buckingham Pi Theorem PowerPoint deck (python-pptx).
- `am-marine-review/` — LaTeX review manuscript: *Sustainable Marine Biofouling Control* (IEEEtran, modular sections, tables, figures).
- `am-hea-marine-review/` — LaTeX review manuscript: *Additive Manufacturing of High-Entropy Alloys for Marine Applications* (IEEEtran, modular sections, tables, figures).

### Manuscript module convention

Each review is a self-contained folder with `main.tex`, modular text in `sections/`, tables in `tables/`, text-based figures in `figures/` (no external graphics dependencies), and a local `references.bib`. Compile from the module folder with `pdflatex main && bibtex main && pdflatex main && pdflatex main`. The `am-hea-marine-review/data/` folder holds machine-readable datasets (`electrochem_quant.csv`: electrochemical synthesis with source identifiers and reference-electrode normalization notes).
