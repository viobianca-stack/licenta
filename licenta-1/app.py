import os
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from transformers import RobertaTokenizer, RobertaForSequenceClassification
import torch
import re
import pickle
import nltk
from nltk.corpus import stopwords
import spacy
import csv
import io

app = Flask(__name__, static_folder='.')
CORS(app)

# ── 1. ÎNCĂRCARE MODELE ──────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

nlp = None
try:
    nlp = spacy.load("en_core_web_sm")
    print("✓ SpaCy încărcat (NER activat)")
except Exception as e:
    print(f"! SpaCy nu s-a încărcat ({e}). Rulează: python -m spacy download en_core_web_sm")

model_path = os.path.join(BASE_DIR, "best_model")
print(f"Se încarcă RoBERTa din: {model_path}")
try:
    tokenizer = RobertaTokenizer.from_pretrained(model_path)
    model = RobertaForSequenceClassification.from_pretrained(model_path)
    model.eval()
    device = torch.device('cuda' if torch.cuda.is_available() else 'cpu')
    model.to(device)
    print(f"✓ RoBERTa încărcat pe {device}")
except Exception as e:
    print(f"✗ Eroare RoBERTa: {e}")
    raise

nb_model, lr_model, tfidf_vec = None, None, None
for fname, attr in [('nb_model.pkl', 'nb'), ('lr_model.pkl', 'lr'), ('tfidf_vec.pkl', 'tfidf')]:
    fpath = os.path.join(BASE_DIR, fname)
    if os.path.exists(fpath):
        with open(fpath, 'rb') as f:
            obj = pickle.load(f)
        if attr == 'nb':      nb_model  = obj
        elif attr == 'lr':    lr_model  = obj
        elif attr == 'tfidf': tfidf_vec = obj
        print(f"✓ {fname} încărcat")

# ── 2. LOGICĂ DE PROCESARE ───────────────────────────────────
nltk.download('stopwords', quiet=True)
STOP = set(stopwords.words('english'))

def clean_nb(text):
    text = str(text).lower()
    text = re.sub(r'http\S+|www\S+', '', text)
    text = re.sub(r'@\w+|#\w+', '', text)
    text = re.sub(r'[^a-z\s]', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    tokens = [w for w in text.split() if w not in STOP and len(w) > 2]
    return " ".join(tokens)

def clean_for_roberta(text):
    text = str(text)
    text = re.sub(r'http\S+|www\S+', '[URL]', text)
    text = re.sub(r'@\w+', '[USER]', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def predict_single(text):
    """Returnează prob_ai pentru un text — folosit intern."""
    text_rb = clean_for_roberta(text)
    enc = tokenizer(text_rb, return_tensors="pt",
                    truncation=True, max_length=128).to(device)
    with torch.no_grad():
        p_ai = torch.softmax(model(**enc).logits, dim=1)[0][1].item()
    return round(p_ai * 100, 2)

def analyze_entities(text):
    if nlp is None:
        return {"entities_found": [], "entity_count": 0, "method": "NER indisponibil"}
    doc = nlp(text)
    entities = [{"text": ent.text, "label": ent.label_} for ent in doc.ents]
    return {"entities_found": entities, "entity_count": len(entities),
            "method": "Named Entity Recognition (SpaCy)"}

def word_importance_analysis(text, n_top=15):
    words = text.split()
    if len(words) < 3:
        return []

    def get_prob_ai(t):
        enc = tokenizer(t, return_tensors="pt",
                        truncation=True, max_length=128).to(device)
        with torch.no_grad():
            return torch.softmax(model(**enc).logits, dim=1)[0][1].item()

    base_prob = get_prob_ai(text)
    seen, scores = set(), {}
    for w in list(dict.fromkeys(words)):
        clean_w = re.sub(r'[^a-zA-Z]', '', w).lower()
        if clean_w in seen or len(clean_w) < 3 or clean_w in STOP:
            continue
        seen.add(clean_w)
        masked = " ".join(x for x in words if x != w)
        if masked.strip():
            scores[w] = base_prob - get_prob_ai(masked)

    top = sorted(scores.items(), key=lambda x: abs(x[1]), reverse=True)[:n_top]
    return [{"word": w, "impact": round(s, 4),
             "direction": "AI" if s > 0 else "UMAN"} for w, s in top]

def get_classic_importance(text, model_type='nb'):
    if not tfidf_vec:
        return []
    active_model = nb_model if model_type == 'nb' else lr_model
    if not active_model:
        return []
    cleaned_text = clean_nb(text)
    words = list(set(cleaned_text.split()))
    if not words:
        return []
    vocab = tfidf_vec.vocabulary_
    word_scores = []
    for word in words:
        if word in vocab:
            idx = vocab[word]
            if model_type == 'nb':
                score = active_model.feature_log_prob_[1][idx] - \
                        active_model.feature_log_prob_[0][idx]
            else:
                score = active_model.coef_[0][idx]
            word_scores.append({
                "word": word,
                "impact": round(float(score), 4),
                "direction": "AI" if score > 0 else "UMAN"
            })
    return sorted(word_scores, key=lambda x: abs(x['impact']), reverse=True)[:15]

def split_sentences(text):
    """Împarte textul în propoziții simple."""
    # Folosim SpaCy dacă e disponibil, altfel regex simplu
    if nlp is not None:
        doc = nlp(text)
        sentences = [sent.text.strip() for sent in doc.sents if len(sent.text.strip()) > 5]
    else:
        sentences = re.split(r'(?<=[.!?])\s+', text)
        sentences = [s.strip() for s in sentences if len(s.strip()) > 5]
    return sentences

def analyze_sentences_classic(text, model_type='nb'):
    """
    Analiză per propoziție pentru Naive Bayes sau Logistic Regression.
    """
    if not tfidf_vec:
        return []
    active_model = nb_model if model_type == 'nb' else lr_model
    if not active_model:
        return []

    sentences = split_sentences(text)
    if not sentences:
        return []

    results = []
    for sent in sentences:
        if len(sent.split()) < 3:
            results.append({
                "text": sent,
                "prob_ai": 50.0,
                "prediction": "INCERT",
                "level": "neutral"
            })
            continue

        cleaned = clean_nb(sent)
        if not cleaned:
            results.append({
                "text": sent,
                "prob_ai": 50.0,
                "prediction": "INCERT",
                "level": "neutral"
            })
            continue

        vec = tfidf_vec.transform([cleaned])
        prob_ai = round(active_model.predict_proba(vec)[0][1] * 100, 2)

        if prob_ai >= 75:
            level = "strong_ai"
        elif prob_ai >= 55:
            level = "weak_ai"
        elif prob_ai <= 25:
            level = "strong_human"
        elif prob_ai <= 45:
            level = "weak_human"
        else:
            level = "neutral"

        results.append({
            "text": sent,
            "prob_ai": prob_ai,
            "prediction": "AI" if prob_ai >= 50 else "UMAN",
            "level": level
        })

    return results


def analyze_sentences(text):
    """
    Analiză per propoziție: fiecare propoziție primește un scor AI individual.
    Returnează lista de propoziții cu scorurile lor.
    """
    sentences = split_sentences(text)
    if not sentences:
        return []

    results = []
    for sent in sentences:
        if len(sent.split()) < 3:
            # Propoziție prea scurtă — scor neutru
            results.append({
                "text": sent,
                "prob_ai": 50.0,
                "prediction": "INCERT",
                "level": "neutral"
            })
            continue

        prob_ai = predict_single(sent)

        # Nivel de certitudine pentru colorare
        if prob_ai >= 75:
            level = "strong_ai"
        elif prob_ai >= 55:
            level = "weak_ai"
        elif prob_ai <= 25:
            level = "strong_human"
        elif prob_ai <= 45:
            level = "weak_human"
        else:
            level = "neutral"

        results.append({
            "text": sent,
            "prob_ai": prob_ai,
            "prediction": "AI" if prob_ai >= 50 else "UMAN",
            "level": level
        })

    return results

# ── 3. RUTE API ──────────────────────────────────────────────

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')


@app.route('/predict', methods=['POST'])
def predict():
    data = request.get_json()
    if not data or 'text' not in data:
        return jsonify({'error': 'Lipsește textul.'}), 400

    text = data['text'].strip()
    if len(text) < 10:
        return jsonify({'error': 'Text prea scurt.'}), 400

    try:
        # A. PROCESARE ROBERTA (text integral)
        text_rb = clean_for_roberta(text)
        enc = tokenizer(text_rb, return_tensors="pt",
                        truncation=True, max_length=128).to(device)
        with torch.no_grad():
            p_ai = torch.softmax(model(**enc).logits, dim=1)[0][1].item()

        # B. RAPORT DE BAZĂ
        out = {
            'prediction': 'AI' if p_ai >= 0.5 else 'UMAN',
            'prob_ai': round(p_ai * 100, 2),
            'prob_human': round((1 - p_ai) * 100, 2),
            'confidence': round(max(p_ai, 1 - p_ai) * 100, 2),
            'ner_analysis': analyze_entities(text),
            'word_importance': word_importance_analysis(text_rb, n_top=15),
            'sentence_analysis': analyze_sentences(text),  # NOU
            'compare': None
        }

        # C. COMPARAȚIE CU MODELE CLASICE
        if tfidf_vec and nb_model and lr_model:
            cleaned_classic = clean_nb(text)
            vec = tfidf_vec.transform([cleaned_classic])
            nb_p = nb_model.predict_proba(vec)[0][1]
            lr_p = lr_model.predict_proba(vec)[0][1]
            out['compare'] = {
                'naive_bayes': {
                    'prob_ai': round(nb_p * 100, 2),
                    'pred': 'AI' if nb_p >= 0.5 else 'UMAN',
                    'importance': get_classic_importance(text, 'nb'),
                    'sentence_analysis': analyze_sentences_classic(text, 'nb')
                },
                'logistic_regression': {
                    'prob_ai': round(lr_p * 100, 2),
                    'pred': 'AI' if lr_p >= 0.5 else 'UMAN',
                    'importance': get_classic_importance(text, 'lr'),
                    'sentence_analysis': analyze_sentences_classic(text, 'lr')
                },
                'roberta': {
                    'prob_ai': out['prob_ai'],
                    'pred': out['prediction'],
                    'importance': out['word_importance'],
                    'sentence_analysis': out['sentence_analysis']
                }
            }

        return jsonify(out)

    except Exception as e:
        print(f"Eroare procesare: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/batch', methods=['POST'])
def batch():
    """
    Analiză batch: primește un fișier CSV cu o coloană 'text'
    și returnează rezultatele pentru fiecare rând.
    Coloana text poate fi numită: text, Text, content, Content, post, Post.
    """
    if 'file' not in request.files:
        return jsonify({'error': 'Lipsește fișierul CSV.'}), 400

    file = request.files['file']
    if not file.filename.endswith('.csv'):
        return jsonify({'error': 'Fișierul trebuie să fie CSV.'}), 400

    try:
        content = file.read().decode('utf-8', errors='replace')
        reader = csv.DictReader(io.StringIO(content))

        # Detectăm coloana cu textul
        text_col = None
        for candidate in ['text', 'Text', 'content', 'Content', 'post', 'Post', 'message']:
            if candidate in (reader.fieldnames or []):
                text_col = candidate
                break

        if text_col is None:
            return jsonify({
                'error': f'Coloana text nu a fost găsită. Coloane disponibile: {reader.fieldnames}'
            }), 400

        results = []
        ai_count = 0
        human_count = 0
        total_prob = 0

        for i, row in enumerate(reader):
            if i >= 500:  # Limităm la 500 rânduri pentru performanță
                break

            text = str(row.get(text_col, '')).strip()
            if len(text) < 5:
                results.append({
                    'index': i + 1,
                    'text_preview': text[:80],
                    'prediction': 'SKIP',
                    'prob_ai': None,
                    'error': 'Text prea scurt'
                })
                continue

            try:
                prob_ai = predict_single(text)
                prediction = 'AI' if prob_ai >= 50 else 'UMAN'

                if prediction == 'AI':
                    ai_count += 1
                else:
                    human_count += 1
                total_prob += prob_ai

                results.append({
                    'index': i + 1,
                    'text_preview': text[:80] + ('...' if len(text) > 80 else ''),
                    'prediction': prediction,
                    'prob_ai': prob_ai,
                    'prob_human': round(100 - prob_ai, 2)
                })
            except Exception as ex:
                results.append({
                    'index': i + 1,
                    'text_preview': text[:80],
                    'prediction': 'EROARE',
                    'prob_ai': None,
                    'error': str(ex)
                })

        valid_count = ai_count + human_count
        summary = {
            'total': len(results),
            'valid': valid_count,
            'ai_count': ai_count,
            'human_count': human_count,
            'ai_percent': round(ai_count / valid_count * 100, 1) if valid_count > 0 else 0,
            'human_percent': round(human_count / valid_count * 100, 1) if valid_count > 0 else 0,
            'avg_prob_ai': round(total_prob / valid_count, 2) if valid_count > 0 else 0
        }

        return jsonify({'summary': summary, 'results': results})

    except Exception as e:
        print(f"Eroare batch: {e}")
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    print("\n🔥 Detector pornit pe http://localhost:5000")
    app.run(debug=False, port=5000)