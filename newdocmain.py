from tkinter import Tk, filedialog, Label, Button, Frame, Text, Scrollbar, VERTICAL, END
import pandas as pd
import numpy as np
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import accuracy_score, classification_report
from sklearn.pipeline import Pipeline
import pickle
import os
import PyPDF2
import warnings
from joblib import parallel_backend
from docx import Document
import win32com.client

# Suppress UserWarnings from joblib
warnings.filterwarnings("ignore", category=UserWarning, module='joblib')

from nltk.stem import WordNetLemmatizer

# === Text Preprocessing Function === #
def preprocess_text(text):
    lemmatizer = WordNetLemmatizer()
    stop_words = set(stopwords.words("english"))
    tokens = word_tokenize(text.lower())
    filtered_text = [lemmatizer.lemmatize(word) for word in tokens if word.isalnum() and word not in stop_words]
    return ' '.join(filtered_text)

# Load and preprocess the dataset
df_train = pd.read_csv('dataset/train.csv')
df_train['processed_text'] = df_train['Description'].apply(preprocess_text)

# Split the training dataset into training and validation sets
X_train, X_val, y_train, y_val = train_test_split(df_train['processed_text'], df_train['Class Index'], test_size=0.2, random_state=42)

# Create a pipeline with TF-IDF Vectorizer and Logistic Regression
pipeline = Pipeline([
    ('tfidf', TfidfVectorizer(max_features=10000, ngram_range=(1, 2))),  # Unigrams and Bigrams
    ('clf', LogisticRegression(max_iter=1000))  # Logistic Regression for better accuracy
])

# Define the grid search parameters for Logistic Regression
param_grid = {
    'tfidf__max_features': [5000, 10000],
    'tfidf__ngram_range': [(1, 1), (1, 2)],
    'clf__C': [0.1, 1.0, 10.0]  # Regularization strength
}

# Perform grid search cross-validation
with parallel_backend('threading'):
    grid_search = GridSearchCV(pipeline, param_grid, cv=5, n_jobs=-1)
    grid_search.fit(X_train, y_train)

# Evaluate the model on the validation set
val_predictions = grid_search.predict(X_val)
accuracy = accuracy_score(y_val, val_predictions)
report = classification_report(y_val, val_predictions)

# Print the results to the terminal
print(f"Best Cross-Validated Accuracy: {grid_search.best_score_:.4f}")
print(f"Validation Accuracy: {accuracy:.4f}")
print("Classification Report:")
print(report)

# Save the best model
best_model = grid_search.best_estimator_
with open('save/model.pkl', 'wb') as model_file:
    pickle.dump(best_model, model_file)

# === GUI for File Input and Class Prediction === #

# File handling and text extraction
uploaded_text = ""
predicted_class = ""
summary_text = ""

def open_file():
    global uploaded_text
    file_path = filedialog.askopenfilename()
    if file_path:
        uploaded_text = extract_text(file_path)

def extract_text(file_path):
    file_extension = os.path.splitext(file_path)[-1].lower()
    text = ""

    if file_extension == '.csv':
        df = pd.read_csv(file_path)
        if 'Description' in df.columns:
            text = " ".join(df['Description'].dropna().tolist())
        else:
            return "No 'Description' column found in the CSV."
    elif file_extension == '.txt':
        with open(file_path, 'r', encoding='utf-8') as file:
            text = file.read()
    elif file_extension == '.pdf':
        with open(file_path, 'rb') as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            for page in reader.pages:
                text += page.extract_text() or ""
    elif file_extension == '.docx':
        doc = Document(file_path)
        text = "\n".join([para.text for para in doc.paragraphs])
    elif file_extension == '.doc':
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(file_path)
        text = doc.Content.Text
        doc.Close()
        word.Quit()
    elif file_extension == '.rtf':
        with open(file_path, 'r', encoding='utf-8') as file:
            text = file.read()  # Simple read for RTF; you might need to process it further.
    elif file_extension == '.msg':
        import extract_msg # type: ignore
        msg = extract_msg.Message(file_path)
        text = msg.body
    else:
        return "Unsupported file format."

    return text if text else "No text extracted or unsupported file format."


def predict_class():
    global predicted_class
    if uploaded_text:
        with open('save/model.pkl', 'rb') as model_file:
            loaded_model = pickle.load(model_file)
        processed_text = preprocess_text(uploaded_text)
        prediction = loaded_model.predict([processed_text])
        class_dict = {1: 'Politics', 2: 'Sports', 3: 'Business/Finance', 4: 'Sci/Tech'}
        predicted_class = class_dict.get(prediction[0], 'Unknown')
        result_label.config(text=f"Predicted Label: {predicted_class}")

def summarize_text():
    global summary_text
    if uploaded_text:
        sentences = sent_tokenize(uploaded_text)
        summary_sentences = sorted(sentences, key=len, reverse=True)[:3] 
        summary_text = ' '.join(summary_sentences)
        summary_text_area.delete(1.0, END)  # Clear previous summary
        summary_text_area.insert(END, summary_text)  # Insert new summary

# Create the GUI
root = Tk()
root.title("Document Classifier")
root.geometry("900x700")  # Set window size
root.configure(bg='lightgray')  # Background color

# Frame for layout
frame = Frame(root, bg='lightgray')
frame.pack(pady=20)

# Label and Button for file input
label = Label(frame, text="Upload a file to classify:", bg='lightgray', font=('Arial', 20))
label.pack(pady=20)

upload_button = Button(frame, text="Upload File", command=open_file, bg='lightblue', font=('Arial', 18))
upload_button.pack(pady=10)

# Frame for buttons to place them side by side
button_frame = Frame(frame, bg='lightgray')
button_frame.pack(pady=10)

# Buttons for Prediction and Summary
predict_button = Button(button_frame, text="Get Predicted Category", command=predict_class, bg='lightgreen', font=('Arial', 15))
predict_button.pack(side='left', padx=10)  # Side by side

summary_button = Button(button_frame, text="Get Summary", command=summarize_text, bg='lightyellow', font=('Arial', 15))
summary_button.pack(side='left', padx=10)  # Side by side

# Labels to display prediction
result_label = Label(frame, text="", bg='lightgray', font=('Arial', 14))
result_label.pack(pady=20)

# Scrollable text area for summary
summary_frame = Frame(frame)
summary_frame.pack(pady=20)

summary_text_area = Text(summary_frame, height=70, width=80, wrap='word')
summary_text_area.pack(side='left')

scrollbar = Scrollbar(summary_frame, command=summary_text_area.yview, orient=VERTICAL)
scrollbar.pack(side='right', fill='y')

summary_text_area.config(yscrollcommand=scrollbar.set)

# Start the GUI
root.mainloop()
