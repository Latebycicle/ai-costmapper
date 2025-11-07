# ai-costmapper
A contextually aware AI cost mapper that allows messy budgeting costs to be mapped by using fuzzy word matching and ai models via api calls to Ollama/Gemini

A Python script to clean, map, and triage messy financial data in Excel. It uses a "Hybrid Triage" model (Exact, Fuzzy, and AI matching) to assist with manual data classification, turning "Lappps (4); DIL Centre 1" into "Capex - Equipment" so a human doesn't have to.



## The Problem

Financial tracking in Excel is powerful, but "Cost Head" or "GL Code" columns are often filled with manually-entered, inconsistent data. This includes:
* Typos: `Asesment and Certifcation`
* Ambiguous entries: `printing for marketing`
* Garbage data: `Lappps (4); DIL Centre 1`

Cleaning this by hand, row-by-row, is a significant bottleneck.

## The Solution

This script acts as an intelligent assistant for the "human-in-the-loop." It doesn't overwrite data. Instead, it reads the messy sheet, applies its logic, and generates a **new `result.xlsx` file**.

This new file contains all the original data plus three new columns:
* `Predicted cost head`
* `Prediction Confidence` (A 0.0-1.0 score)
* `AI Confidence` (A 1-10 score, only for AI-powered guesses)

The output `messy` sheet is also formatted for easy review, with 3-color (Red-Yellow-Green) gradients on the confidence columns.

## How It Works: The Hybrid TTriage

The script attacks the problem in three steps, in order of cost and confidence. It only escalates a row to the next step if the current one fails.

1.  **Step 1: Exact Match (Confidence: 100%)**
    * **Logic:** Does the messy string `Capex - Equipment` exactly match a valid `actual GL Head`?
    * **Result:** A perfect, 100% confidence match.

2.  **Step 2: Fuzzy Match (Confidence: 90%+)**
    * **Logic:** Does the messy string `Asesment and Certifcation` look *really similar* to a valid GL Head?
    * **Result:** Catches obvious typos with high confidence (e.g., 92%).

3.  **Step 3: AI Escalation (Confidence: 10-90%)**
    * **Logic:** For the true garbage that failed Steps 1 & 2. The script builds a **one-time system prompt** containing the *entire `rules` sheet* (all valid heads and all rule examples).
    * It then sends this context, plus the one messy string, to a local Ollama model (`qwen3:4b`).
    * **Result:** The AI, now primed with your team's specific business logic, can correctly guess that "printing for marketing" belongs to "Operational Marketing" and not "Consumables - Printing."

## 📁 Project Structure

```
fin-cost-triage/ │ ├── Data/ <-- MUST BE IN .gitignore. Contains your sensitive data. │ └── Test.xlsx <-- Your input file with 'rules' and 'messy' sheets │ ├── .gitignore <-- Crucial! Make sure 'Data/' is in here. ├── result.xlsx <-- This is the final, formatted output file ├── triage.py <-- The main Python script └── requirements.txt <-- Project dependencies
```
## 🚀 Setup & Installation

1.  **Clone the Repository:**
    ```bash
    git clone [https://github.com/your-username/fin-cost-triage.git](https://github.com/your-username/fin-cost-triage.git)
    cd fin-cost-triage
    ```
2.  **IMPORTANT: Configure `.gitignore`**
    Make sure your `.gitignore` file contains a line for `Data/` to prevent your proprietary Excel files from ever being committed.
    ```
    # .gitignore
    Data/
    venv/
    *.pyc
    ```

3.  **Create a Virtual Environment** (Recommended):
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

4.  **Install Dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

5.  **Install and Run Ollama:**
    This script *requires* [Ollama](https.ollama.com) to be installed and running locally.
    ```bash
    # After installing Ollama, pull the model specified in the script
    ollama pull qwen3:4b
    ```

## ▶️ How to Run

1.  Ensure your source file is located at `Data/Test.xlsx`.
2.  Ensure your `rules` and `messy` sheets are correctly named inside the Excel file.
3.  **Make sure the Ollama application is running in the background.**
4.  Run the script:
    ```bash
    python triage.py
    ```

The script will print its progress for each step and notify you when `result.xlsx` is complete.

## 📝 `requirements.txt`

```
pandas openpyxl ollama thefuzz numpy tqdm
```
## 🔧 Configuration

To change key parameters, edit the constants at the top of `triage.py`:

* **File Paths:** `INPUT_FILE` and `OUTPUT_FILE`.
* **Ollama Model:** `OLLAMA_MODEL` (e.g., `'qwen3:4b'`).
* **Fuzzy Match Sensitivity:** `FUZZY_THRESHOLD` (default: `90`).