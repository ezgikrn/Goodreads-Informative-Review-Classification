# MEB Goodreads Informative Review Classification

This repository contains the code and shared datasets used to classify Turkish Goodreads book reviews as informative or non-informative.

## Overview

The project includes:

- Goodreads review collection and export utilities
- preprocessing and dataset validation scripts
- train/test dataset preparation
- classical machine learning experiments for informative review classification
- optional LLM-based zero-shot classification experiments
- raw dataset files prepared for sharing and reproducibility

## Main Files

- `01_MEB_Book_List_Scraper.py`: scrapes the MEB book list
- `02_Goodreads_Turkish_Review_Scraper.py`: collects Turkish Goodreads reviews
- `02A_Goodreads_TR_Export_Helper.py`: helper module for Goodreads export logic
- `03_Preprocessing_Clean_Data_Check.py`: checks whether preprocessing outputs and required columns are ready
- `04_Prepare_Informative_Dataset.py`: prepares the train/test dataset structure
- `05_Informative_ML_Model_Comparison.py`: runs machine learning experiments
- `06_Informative_LLM_Experiments.py`: runs zero-shot LLM experiments
- `unlabeled_raw_embedding_corpus.xlsx`: raw unlabeled review corpus with the `review_text` column, corresponding to the external review pool used for the Word2Vec and FastText experiments
- `labeled_raw_ml_dataset.xlsx`: raw labeled dataset with the `sample_id`, `review_text`, and `informative_label` columns

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

Prepare the hold-out train/test structure and assign training folds:

```bash
python 04_Prepare_Informative_Dataset.py
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
- The repository includes the code and shared dataset files used in the experimental pipeline.
- The shared dataset files contain raw review text and do not include platform-specific review identifiers.
