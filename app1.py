import os
import msal 
import spacy
import torch
from flask import Flask, render_template, redirect, url_for, session, request, flash
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
import PyPDF2 
import docx 
from sentence_transformers import SentenceTransformer
from datetime import timedelta

# Initialize global dictionary for processed documents
preprocessed_docs = {}

# Initialize Flask and Flask-Login
app = Flask(__name__, template_folder='templates', static_folder='static')
app.secret_key = os.getenv("SECRET_KEY", "bR7Fj9PzS2Xv5Lq8Mn6Yt3Wk9Qd0Jx4Z")  # Secure in production
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(days=1)

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"

# Office 365 App credentials (use environment variables in production)
CLIENT_ID = os.getenv("CLIENT_ID", "04ca7d78-4284-4358-83f7-53cfbaf720e0")
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "sKM8Q~dkI1HjNLZif-w.sZwVv1PvBZt85mrIqawv")
AUTHORITY = "https://login.microsoftonline.com/1f8e1b2e-5d9b-44a7-91fc-dbe7d9a51b15"
REDIRECT_URI = "http://localhost:5000/login/callback"
SCOPE = ["User.Read"]

@app.route('/')
def home():
    if current_user.is_authenticated:
        return redirect(url_for('chat'))
    return redirect(url_for('login'))

# Load NLP model
nlp = spacy.load("en_core_web_sm")
model = SentenceTransformer("all-MiniLM-L6-v2")

# Function to capture user session
def capture_user_session(user_info):
    session.permanent = True
    session["user"] = user_info
    session["email"] = user_info.get("preferred_username")
    login_user(User(session["email"], session["email"]))
    print(f"✅ User session captured: {session['email']}")

# 🔹 Extract text from documents
def extract_text_from_pdf(file_path):
    text = ""
    try:
        with open(file_path, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                page_text = page.extract_text() or ""
                text += page_text.strip() + " "
    except Exception as e:
        print(f"❌ Error reading {file_path}: {e}")
    return text

def extract_text_from_docx(file_path):
    text = ""
    try:
        doc = docx.Document(file_path)
        text = "\n".join([para.text.strip() for para in doc.paragraphs if para.text.strip()])
    except Exception as e:
        print(f"❌ Error reading {file_path}: {e}")
    return text

# 🔹 Process documents at startup
def process_documents():
    global preprocessed_docs
    folder = "documents"
    
    if not os.path.exists(folder):
        print("❌ Documents folder not found!")
        return

    preprocessed_docs.clear()  # Ensure no duplicate entries

    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        text = ""

        if filename.endswith('.pdf'):
            text = extract_text_from_pdf(file_path)
        elif filename.endswith('.docx'):
            text = extract_text_from_docx(file_path)
        else:
            continue

        if text:
            sentences = text.lower().split('. ')
            preprocessed_docs[filename] = [sentence.strip() for sentence in sentences]

    print(f"✅ Loaded {len(preprocessed_docs)} documents: {list(preprocessed_docs.keys())}")

process_documents()

# User class (Flask-Login)
class User(UserMixin):
    def __init__(self, id, email):
        self.id = id
        self.email = email

# Load user from session
@login_manager.user_loader
def load_user(user_id):
    email = session.get("email")
    return User(user_id, email) if email else None

# 🔹 Office 365 Login
@app.route('/login')
def login():
    if current_user.is_authenticated:
        return redirect(url_for('chat'))
    
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    auth_url = msal_app.get_authorization_request_url(SCOPE, redirect_uri=REDIRECT_URI, response_mode='form_post')
    return redirect(auth_url)

# 🔹 Authentication Callback
@app.route('/login/callback')
def login_callback():
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    code = request.args.get('code')
    
    if not code:
        return "Authorization failed: No code received", 400  # Handle missing code error

    result = msal_app.acquire_token_by_authorization_code(code, scopes=SCOPE, redirect_uri=REDIRECT_URI)

    if "access_token" in result:
        user_info = result.get("id_token_claims")  # Extract user info
        if not user_info:
            return "Authentication failed", 400
        
        session["user"] = user_info
        session["email"] = user_info.get("preferred_username", "unknown_user@example.com")  # Safer handling
        user = User(user_info["preferred_username"], user_info["preferred_username"])
        login_user(user)

        return redirect(url_for('chat'))
    else:
        return "Login failed: " + str(result.get("error_description", "Unknown error")), 400


@app.route('/chat', methods=['GET', 'POST'])
@login_required  # Ensures only logged-in users can access chat
def chat():
    print("Current User:", current_user)  # Debugging
    response = None
    if request.method == 'POST':
        user_input = request.form['user_input']
        response = get_answer_from_documents(user_input)
        print("User Input:", user_input)  # Debugging
        print("Response:", response)
    return render_template('chat.html', response=response)

def get_answer_from_documents(query):
    query_embedding = model.encode(query, convert_to_tensor=True)
    ranked_sentences = []

    print(f"🔍 Processing query: {query}")
    print(f"📂 Available documents: {preprocessed_docs.keys()}")

    for doc_name, sentences in preprocessed_docs.items():
        for sentence in sentences:
            sentence_embedding = model.encode(sentence, convert_to_tensor=True)
            similarity_score = torch.nn.functional.cosine_similarity(query_embedding, sentence_embedding, dim=0).item()
            ranked_sentences.append((similarity_score, sentence))

    # Sort results by highest similarity score
    ranked_sentences.sort(reverse=True, key=lambda x: x[0])

    # Select top 3 sentences for a structured response
    top_sentences = [s[1] for s in ranked_sentences[:3] if s[0] > 0.5]  # Only take relevant ones

    if top_sentences:
        structured_response = "Here’s what I found:\n\n" + "\n".join(top_sentences)  # New lines preserved
        print(f"✅ Structured Response:\n{structured_response}")
        return structured_response

    print("❌ No relevant information found.")
    return "I'm sorry, but I couldn't find any relevant information in the documents."



# 🔹 Logout Route
@app.route('/logout')
@login_required
def logout():
    logout_user()
    session.clear()

    # Redirect user to Microsoft logout URL
    logout_url = f"https://login.microsoftonline.com/1f8e1b2e-5d9b-44a7-91fc-dbe7d9a51b15/oauth2/v2.0/logout?post_logout_redirect_uri=http://localhost:5000/"
    return redirect(logout_url)


# 🔹 Run Flask App
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
