# MEB Goodreads Informative Review Classification

This repository contains the code and workbook used to classify Turkish Goodreads book reviews as informative or non-informative.

## Overview

The project includes:

- Goodreads review collection and export utilities
- preprocessing and dataset validation scripts
- train/test dataset preparation
- classical machine learning experiments for informative review classification
- optional LLM-based zero-shot classification experiments

## Main Files

- `01_MEB_Book_List_Scraper.py`: scrapes the MEB book list
- `02_Goodreads_Turkish_Review_Scraper.py`: collects Turkish Goodreads reviews
- `02A_Goodreads_TR_Export_Helper.py`: helper module for Goodreads export logic
- `03_Preprocessing_Clean_Data_Check.py`: checks whether preprocessing outputs and required columns are ready
- `04_Prepare_Informative_TrainTest_10Fold.py`: prepares the train/test dataset structure
- `05_Informative_ML_Model_Comparison.py`: runs machine learning experiments
- `06_Informative_LLM_Experiments.py`: runs zero-shot LLM experiments
- `MEB_100_Goodreads_ML_LLM.xlsx`: main workbook used by the scripts

## Experimental Setting

- Total manually labeled reviews: `1000`
- Train set: `800`
- Held-out test set: `200`
- A `hold-out train/test` split is used, and `10-fold GridSearchCV` is applied during training for model selection.

## Requirements

The scripts use Python and depend mainly on:

- `openpyxl`
- `scikit-learn`
- `gensim`
- `numpy`
- `transformers`
- `torch`

Some scripts may also require API access for LLM experiments.

## Environment Variables

If you run the LLM script, create a local `.env` file and define only the keys you need:

```env
GROQ_API_KEY=
MISTRAL_API_KEY=
GEMINI_API_KEY=
GOODREADS_ENV_PATH=
```

Do not commit `.env` to GitHub.

## Basic Usage

Validate preprocessing outputs:

```bash
python 03_Preprocessing_Clean_Data_Check.py
```

Prepare the V4 hold-out train/test structure and assign training folds:

```bash
python 04_Prepare_Informative_TrainTest_10Fold.py
```

Run machine learning experiments:

```bash
python 05_Informative_ML_Model_Comparison.py --mode grid_all
```

Run LLM experiments:

```bash
python 06_Informative_LLM_Experiments.py
```

## Notes

- `__pycache__`, `.env`, and temporary files should not be committed.
- The LLM experiments depend on external APIs and may not be fully reproducible over time.
- The workbook is included because it is part of the experimental pipeline.
