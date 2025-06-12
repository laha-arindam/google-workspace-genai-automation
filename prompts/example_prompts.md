# 🧠 Example Prompts Used in Google Workspace Automation

This file contains curated examples of prompts used with the OpenAI GPT-4o API for automating tasks in Gmail, Google Sheets, and Google Slides via Google Apps Script.

---

## 📊 Spreadsheet Assistant

### 🔹 Prompt 1: Generate a Google Sheets formula
> "Write a Google Sheets formula to calculate the average sales from column C where the values in column A are 'North'."

### 🔹 Prompt 2: Create a bar chart from spreadsheet data
> "Generate Google Apps Script code to create a bar chart showing revenue from column B using the months listed in column A."

### 🔹 Prompt 3: Detect errors in formulas
> "Find the error in the formula: =IF(A2>B2, 'Yes', 'No)"

---

## 📽️ Slide Deck Generator

### 🔹 Prompt 1: Create a presentation on AI in education
> "Create a 5-slide presentation on 'The Future of Artificial Intelligence in Education.' Include titles and bullet points for each slide."

**Expected GPT Output:**
```json
[
  {
    "title": "Introduction",
    "content": "Definition of AI and its relevance in education."
  },
  {
    "title": "Benefits of AI",
    "content": "Personalized learning, grading automation, data-driven insights."
  },
  ...
]
```

---

### 🔹 Prompt 2: Business pitch deck
> "Generate a pitch deck for a startup that uses drones to deliver groceries in urban areas."

---

## 🧠 Prompt Engineering Strategies Used

- ✅ Clear and specific instructions  
- ✅ Role prompting (e.g., “You are a spreadsheet assistant…”)  
- ✅ Output format specification (JSON for slide content)  
- ✅ Natural language to code translation  

---

## 📎 Usage

All prompts were passed from Google Apps Script using `UrlFetchApp.fetch()`  
calls to OpenAI's Chat Completion API (GPT-4o), with basic string templating or input collected from Google Sheets dialogs.
