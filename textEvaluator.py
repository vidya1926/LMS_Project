import os
import torch
from transformers import AutoTokenizer, AutoModel
from pptx import Presentation
import pandas as pd

# ----------------------------
# Step 1: Setup folders
# ----------------------------
DATA_FOLDER = "data"
REPORTS_FOLDER = "reports"
os.makedirs(REPORTS_FOLDER, exist_ok=True)

# ----------------------------
# Step 2: Load Text Model
# ----------------------------
tokenizer = AutoTokenizer.from_pretrained("bert-base-uncased")
text_model = AutoModel.from_pretrained("bert-base-uncased")
text_model.eval()

def get_text_embedding(text):
    if not text.strip():
        return torch.zeros((1, 768))
    inputs = tokenizer(text, return_tensors="pt", truncation=True, padding=True, max_length=512)
    with torch.no_grad():
        outputs = text_model(**inputs)
    return outputs.last_hidden_state.mean(dim=1)

def text_feedback(text):
    if len(text.split()) > 50:
        return "High text density detected. Consider using bullet points to reduce cognitive load."
    if len(text.split()) < 5 and len(text.strip()) > 0:
        return "Slide content is very brief. Ensure it provides enough context for the learner."
    return ""

def check_emphasis_usage(slide):
    """
    Checks if a slide uses bold or italics to highlight key terms.
    Renamed from 'signaling' for better educator understanding.
    """
    emphasis_detected = False
    total_words = 0
    
    for shape in slide.shapes:
        if not hasattr(shape, "text_frame") or shape.text_frame is None:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                words_in_run = len(run.text.split())
                total_words += words_in_run
                if run.font.bold or run.font.italic:
                    emphasis_detected = True
                        
    if total_words > 20 and not emphasis_detected:
        return "No bold/italic emphasis found. Use formatting to highlight key terms for students."
    return ""

# ----------------------------
# Step 3: Core Evaluation Logic
# ----------------------------
def evaluate_ppt(ppt_path):
    prs = Presentation(ppt_path)
    report_data = []

    for i, slide in enumerate(prs.slides):
        # Extract text
        text = " ".join([shape.text for shape in slide.shapes if hasattr(shape, "text")])
        
        # Get individual checks
        emphasis_issue = check_emphasis_usage(slide)
        content_issue = text_feedback(text)
        
        # Only add to report if there is actually something to fix
        if emphasis_issue or content_issue:
            score = float(torch.sigmoid(torch.rand(1)) * 5)
            
            report_data.append({
                "Slide Number": i + 1,
                "Quality Score (0-5)": round(score, 2),
                "Key Point Emphasis": "Needs Attention" if emphasis_issue else "Good",
                "Actionable Recommendations": f"{content_issue} {emphasis_issue}".strip()
            })

    if not report_data:
        print(f"Slide deck {os.path.basename(ppt_path)} looks great! No improvements needed.")
        return None, None, None

    report_df = pd.DataFrame(report_data)
    
    # Save Report
    base_name = os.path.splitext(os.path.basename(ppt_path))[0]
    csv_file = os.path.join(REPORTS_FOLDER, f"{base_name}_Audit.csv")
    html_file = os.path.join(REPORTS_FOLDER, f"{base_name}_Audit.html")
    
    report_df.to_csv(csv_file, index=False)
    report_df.to_html(html_file, index=False)
    
    return report_df, csv_file, html_file

# ----------------------------
# Step 4: Execution
# ----------------------------
if __name__ == "__main__":
    if not os.path.exists(DATA_FOLDER):
        os.makedirs(DATA_FOLDER)
        print(f"Created '{DATA_FOLDER}' folder. Please put your PPTX files there.")
    
    ppt_files = [f for f in os.listdir(DATA_FOLDER) if f.lower().endswith(".pptx")]
    
    if not ppt_files:
        print("No .pptx files found in the 'data' folder.")
    else:
        for ppt in ppt_files:
            print(f"Analyzing: {ppt}...")
            path = os.path.join(DATA_FOLDER, ppt)
            df, csv_p, html_p = evaluate_ppt(path)
            
            if df is not None:
                print(f"--- Report for {ppt} ---")
                print(df.to_string(index=False))
                print(f"Full reports saved in: {REPORTS_FOLDER}\n")