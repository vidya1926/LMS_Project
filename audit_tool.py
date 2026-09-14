import os
from pydoc import text
import torch
import numpy as np
import matplotlib.pyplot as plt
from transformers import AutoTokenizer, AutoModel
from pptx import Presentation
import pandas as pd

DATA_FOLDER = "data"
REPORTS_FOLDER = "reports"
os.makedirs(REPORTS_FOLDER, exist_ok=True)

print("Loading BERT for Semantic Analysis...")
tokenizer = AutoTokenizer.from_pretrained("bert-base-uncased")
text_model = AutoModel.from_pretrained("bert-base-uncased")
text_model.eval()

def get_semantic_density(text):
    if not text.strip() or len(text.split()) < 5:
        return 0.0
    inputs = tokenizer(text, return_tensors="pt", truncation=True, padding=True, max_length=512)
    with torch.no_grad():
        outputs = text_model(**inputs)
        return float(outputs.last_hidden_state.std().item())
def analyze_slide_pedagogy(slide, text):
    word_count = len(text.split())
    feedback = []
    penalty = 0

    if word_count > 50:
        feedback.append("High Text Density (Cognitive Overload).")
        penalty += 1.5
    elif 0 < word_count < 5:
        feedback.append("Insufficient context for learners.")
        penalty += 0.5

    emphasis_detected = False
    for shape in slide.shapes:
        if hasattr(shape, "text_frame") and shape.text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.bold or run.font.italic:
                        emphasis_detected = True

    if word_count > 20 and not emphasis_detected:
        feedback.append("Signaling Deficiency: Key terms not highlighted.")
        penalty += 1.0

    score = max(0.5, 5.0 - penalty)
    return score, " | ".join(feedback) if feedback else "Optimal Design"

def generate_visuals(summary_df):
    plt.figure(figsize=(10, 6))
    df_sorted = summary_df.sort_values(by="Average_Score", ascending=False)
    plt.bar(df_sorted["File_Name"], df_sorted["Average_Score"], color='skyblue', edgecolor='black')
    plt.axhline(y=3.5, color='red', linestyle='--', label='Quality Threshold')
    plt.title("Pedagogical Quality Score per Slide Deck")
    plt.ylabel("Score (0-5)")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(os.path.join(REPORTS_FOLDER, "audit_results_chart.png"))
    plt.close()
# def run_full_audit():
#     if not os.path.exists(DATA_FOLDER):
#         os.makedirs(DATA_FOLDER)
#         print("Created 'data' folder. Place .pptx files there and rerun.")
#         return

# ppt_files = [f for f in os.listdir(DATA_FOLDER) if f.lower().endswith(".pptx")]
# all_summaries = []

# for ppt in ppt_files:
#     print(f"Analyzing: {ppt}...")
#     path = os.path.join(DATA_FOLDER, ppt)
#     prs = Presentation(path)
#     report_data = []

#     for i, slide in enumerate(prs.slides):
#         text = " ".join([s.text for s in slide.shapes if hasattr(s, "text")])
#         density = get_semantic_density(text)
#         score, advice = analyze_slide_pedagogy(slide, text)
#         report_data.append({"Slide": i+1, "Score": score, "Density": density, "Advice": advice})

#     df = pd.DataFrame(report_data)
#     all_summaries.append({
#         "File": ppt,
#         "Average_Score": round(df["Score"].mean(), 2),
#         "Issues": df[df["Advice"] != "Optimal Design"].shape[0]
#     })
#     df.to_csv(os.path.join(REPORTS_FOLDER, f"{os.path.splitext(ppt)[0]}_Audit.csv"), index=False)

# if all_summaries:
#     summary_df = pd.DataFrame(all_summaries)
#     summary_df.to_csv(os.path.join(REPORTS_FOLDER, "Batch_Summary.csv"), index=False)
#     generate_visuals(summary_df)
#     print("Success! Check the 'reports' folder.")
# if __name__ == "__main__":
#     run_full_audit()

# Step 5: Main Execution
def run_full_audit():
    if not os.path.exists(DATA_FOLDER):
        os.makedirs(DATA_FOLDER)
        print("Created 'data' folder. Place .pptx files there and rerun.")
        return

    ppt_files = [f for f in os.listdir(DATA_FOLDER) if f.lower().endswith(".pptx")]
    
    # We will also keep a list for the Batch Summary (Macro-level)
    all_summaries = []

    for ppt in ppt_files:
        print(f"Generating Micro-Level Audit for: {ppt}...")
        path = os.path.join(DATA_FOLDER, ppt)
        prs = Presentation(path)
        report_data = []

        for i, slide in enumerate(prs.slides):
            text = " ".join([s.text for s in slide.shapes if hasattr(s, "text")])
            
            # Semantic and Pedagogical checks
            density = get_semantic_density(text)
            score, advice = analyze_slide_pedagogy(slide, text)
            
            # Building the Micro-level row
            report_data.append({
                "Slide_Number": i + 1,
                "Quality_Score": score,
                "Semantic_Density": round(density, 4),
                "Instructional_Advice": advice
            })

        # --- SAVE INDIVIDUAL CSV (Micro-Level) ---
        # This creates "Lecture_1_Audit.csv", "Lecture_2_Audit.csv", etc.
        file_base_name = os.path.splitext(ppt)[0]
        micro_csv_path = os.path.join(REPORTS_FOLDER, f"{file_base_name}_Micro_Audit.csv")
        
        df_micro = pd.DataFrame(report_data)
        df_micro.to_csv(micro_csv_path, index=False)
        print(f"Saved micro-level report: {micro_csv_path}")

        # Collect data for the Batch Summary
        all_summaries.append({
            "File_Name": ppt,
            "Average_Score": round(df_micro["Quality_Score"].mean(), 2),
            "Total_Slides": len(df_micro),
            "Issues_Detected": df_micro[df_micro["Instructional_Advice"] != "Optimal Design"].shape[0]
        })

    # Save the Batch Summary (Macro-Level)
    if all_summaries:
        summary_df = pd.DataFrame(all_summaries)
        summary_df.to_csv(os.path.join(REPORTS_FOLDER, "Batch_Executive_Summary.csv"), index=False)
        generate_visuals(summary_df)
        print("\nAll audits complete. Individual CSVs are in the 'reports' folder.")

if __name__ == "__main__":
    run_full_audit()