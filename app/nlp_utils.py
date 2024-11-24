import spacy
from transformers import pipeline

nlp_spacy = spacy.load("en_core_web_sm")
summarizer = pipeline("summarization", model="facebook/bart-large-cnn")

def extract_key_details(text):
    """Extract goals, deliverables, timeline, and evaluation criteria using spaCy."""
    doc = nlp_spacy(text)
    goals, deliverables, timeline, criteria = [], [], [], []

    for ent in doc.ents:
        if ent.label_ in ["ORG", "PRODUCT"]:
            goals.append(ent.text)
        elif ent.label_ == "DATE":
            timeline.append(ent.text)
        elif ent.label_ in ["MONEY", "PERCENT"]:
            deliverables.append(ent.text)
        elif "evaluation" in ent.text.lower():
            criteria.append(ent.text)

    return {
        "Goals": " | ".join(goals) if goals else "No specific goals detected.",
        "Deliverables": " | ".join(deliverables) if deliverables else "No specific deliverables detected.",
        "Timeline": " | ".join(timeline) if timeline else "No timeline detected.",
        "Evaluation Criteria": " | ".join(criteria) if criteria else "No evaluation criteria detected.",
    }

def summarize_text(text, max_length=130, min_length=30):
    """Summarize long text using a Hugging Face model."""
    if len(text.split()) < 50:
        return text
    summary = summarizer(
        text,
        max_length=max_length,
        min_length=min_length,
        do_sample=False,
    )
    return summary[0]["summary_text"]

def generate_contextual_summary(text, context):
    """Use GPT models to generate section-specific summaries."""
    from transformers import pipeline
    summarizer = pipeline("text-generation", model="gpt-3.5-turbo")
    prompt = f"Summarize the following text for {context}: {text}"
    result = summarizer(prompt, max_length=150, min_length=50, do_sample=False)
    return result[0]["generated_text"]

