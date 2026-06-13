# Drafting Chapter 2: Materials and Methods

This plan outlines the approach for drafting **Phase 2: Materials and Methods (Chapter Two)** of the Pharm.D project on the acute oral toxicity and antidysmenorrhea effects of Clove (*Syzygium aromaticum*).

## User Review Required

> [!IMPORTANT]
> Please review this plan to ensure the structure and approach for Chapter 2 align with your expectations. Once approved, I will begin drafting the chapter manually, section by section, adhering strictly to the `no_bullshit.md` rules (no em-dashes, no hyperbole, objective academic tone).

## Open Questions

> [!WARNING]
> 1. **Data specifics:** For the plant collection (Section 2.1.3), should I leave placeholders for the specific dates, locations, and voucher numbers to be filled in later after the actual wet lab work, or do you have this information already?
> 2. **Animal specifics:** For Section 2.1.4, do you have specific weights, ages, or source locations for the mice/rats, or should I use standard protocol placeholders (e.g., "adult female mice weighing 20-25g")?
> 3. **Drafting approach:** Would you prefer I draft Chapter 2 in one single Markdown file (e.g., `01_Chapter_2_Materials_and_Methods.md`) or split it into separate files for Materials, Methods, and Data Analysis, similar to Chapter 1?

## Proposed Changes

We will create the Markdown files for Chapter 2 inside a new directory: `src/project/02_Chapter_2_MATERIALS_AND_METHODS/`.

### Chapter 2 Structure

Based on `roadmap.md`, the following content will be drafted:

#### [NEW] `01_Materials.md`
- **2.1.1 Equipment list:** Organ bath, force transducer, PowerLab, micropipettes, cages, surgical kit, etc.
- **2.1.2 Chemicals and reagents:** NaCl, KCl, CaCl₂, glucose, NaHCO₃, ethanol, oxytocin, estradiol benzoate, tween 80, EDTA, formalin, ibuprofen.
- **2.1.3 Collection and extraction of plant material:** Where purchased, identification/voucher, Soxhlet extraction with hydroethanolic solvent (1:1), yield percentage.
- **2.1.4 Experimental animals:** Species, sex, weight range, source, acclimatization period, housing conditions. (Ethical approval reference has been removed).

#### [NEW] `02_Methods.md`
- **2.2.1 Acute oral toxicity test (OECD 423):** Dose selection (300 mg/kg and 2000 mg/kg step), grouping (n=3 per step), 14-day observation period, parameters monitored, necropsy.
- **2.2.2 Antidysmenorrhea study (oxytocin-induced writhing model):** Estradiol benzoate sensitization, oxytocin challenge, writhing count, group design.
- **2.2.3 Isolated uterine tissue study (organ bath):** Tissue preparation, physiological saline solution, equilibration, cumulative concentration-response to extract on spontaneous, oxytocin-induced, and KCl-induced contractions.

#### [NEW] `03_Data_Analysis.md`
- **2.3 Data Analysis:** Statistical tests used (one-way ANOVA, Dunnett's post-hoc), software (GraphPad Prism), p-value threshold (p < 0.05), data presentation format (mean ± SEM).

## Verification Plan

### Manual Verification
- I will manually ensure no em-dashes (`—`) or hyperbolic adjectives (e.g., "massive", "profound", "astonishing") are used in the text.
- I will verify that the methods are written in the standard past-passive voice expected in scientific literature (e.g., "The animals were acclimatized...", "The extract was prepared...").
- I will compile the draft using the existing Python scripts and present the compiled `.docx` to you for review.
