import os
import re
import json
import shutil
import requests

from dotenv import load_dotenv
from flask import (
    Flask,
    render_template,
    request,
    jsonify,
    Response,
    stream_with_context
)

import pymupdf
from docx import Document

import chromadb
from chromadb.utils.embedding_functions import DefaultEmbeddingFunction


# ============================================================
# LOAD ENVIRONMENT VARIABLES
# ============================================================

load_dotenv()


# ============================================================
# FLASK APP
# ============================================================

app = Flask(__name__)


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
CHROMA_FOLDER = os.path.join(BASE_DIR, "chroma_db")

MAX_FILE_SIZE = 10 * 1024 * 1024

MAX_RESULTS = 10
MAX_CONTEXT_CHUNKS = 4

ALLOWED_EXTENSIONS = {"pdf", "docx"}

NO_ANSWER = (
    "I couldn't find this information in the currently uploaded document."
)


# ============================================================
# OLLAMA CLOUD
# ============================================================

OLLAMA_URL = os.getenv(
    "OLLAMA_URL",
    "https://ollama.com/api/generate"
)

OLLAMA_MODEL = os.getenv(
    "OLLAMA_MODEL",
    "gpt-oss:120b"
)

OLLAMA_API_KEY = os.getenv("OLLAMA_API_KEY")


# ============================================================
# CREATE REQUIRED FOLDERS
# ============================================================

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CHROMA_FOLDER, exist_ok=True)


# ============================================================
# CHROMA DATABASE - LAZY LOADING
# ============================================================

# Important for Render Free:
# Do NOT load the embedding model while the server is starting.
# Load it only when a document actually needs to be processed.

chroma_client = chromadb.PersistentClient(
    path=CHROMA_FOLDER
)

embedding_function = None
collection = None


def get_embedding_function():
    """
    Load Chroma's default embedding function only when needed.
    """

    global embedding_function

    if embedding_function is None:
        print("Loading embedding function...")

        embedding_function = DefaultEmbeddingFunction()

        print("Embedding function loaded.")

    return embedding_function


def get_collection():
    """
    Get the current Chroma collection.
    Creates it if it does not exist.
    """

    global collection

    if collection is not None:
        return collection

    ef = get_embedding_function()

    try:
        collection = chroma_client.get_collection(
            name="current_document",
            embedding_function=ef
        )

        print("Existing Chroma collection loaded.")

    except Exception:

        collection = chroma_client.create_collection(
            name="current_document",
            metadata={
                "hnsw:space": "cosine"
            },
            embedding_function=ef
        )

        print("New Chroma collection created.")

    return collection


# ============================================================
# FILE HELPERS
# ============================================================

def allowed_file(filename):
    """
    Check whether uploaded file is PDF or DOCX.
    """

    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower()
        in ALLOWED_EXTENSIONS
    )


# ============================================================
# TEXT CLEANING
# ============================================================

def clean_text(text):
    """
    Clean extracted document text.
    """

    if not text:
        return ""

    text = text.replace("\x00", " ")

    text = re.sub(
        r"[ \t]+",
        " ",
        text
    )

    text = re.sub(
        r"\n\s*\n\s*\n+",
        "\n\n",
        text
    )

    return text.strip()


# ============================================================
# PDF EXTRACTION
# ============================================================

def extract_pdf(file_path):
    """
    Extract text from PDF page by page.
    """

    pages = []

    try:
        pdf = pymupdf.open(file_path)

        for page_number, page in enumerate(pdf, start=1):

            text = page.get_text("text")

            text = clean_text(text)

            if text:
                pages.append({
                    "page": page_number,
                    "text": text
                })

        pdf.close()

    except Exception as e:
        print("PDF extraction error:", e)
        raise

    return pages


# ============================================================
# DOCX EXTRACTION
# ============================================================

def extract_docx(file_path):
    """
    Extract paragraphs and tables from DOCX.
    """

    document = Document(file_path)

    parts = []

    # --------------------------------------------------------
    # Paragraphs
    # --------------------------------------------------------

    for paragraph in document.paragraphs:

        text = clean_text(paragraph.text)

        if text:
            parts.append(text)

    # --------------------------------------------------------
    # Tables
    # --------------------------------------------------------

    for table in document.tables:

        for row in table.rows:

            cells = []

            for cell in row.cells:

                cell_text = clean_text(cell.text)

                if cell_text:
                    cells.append(cell_text)

            if cells:
                parts.append(" | ".join(cells))

    full_text = "\n".join(parts)

    full_text = clean_text(full_text)

    if not full_text:
        return []

    return [
        {
            "page": 0,
            "text": full_text
        }
    ]


# ============================================================
# DOCUMENT EXTRACTION
# ============================================================

def extract_document(file_path):
    """
    Extract text according to file type.
    """

    extension = file_path.rsplit(".", 1)[1].lower()

    if extension == "pdf":
        return extract_pdf(file_path)

    if extension == "docx":
        return extract_docx(file_path)

    raise ValueError("Unsupported file type.")


# ============================================================
# HEADING DETECTION
# ============================================================

HEADING_KEYWORDS = [
    "introduction",
    "definition",
    "overview",
    "background",
    "objectives",
    "objective",
    "scope",
    "importance",
    "applications",
    "advantages",
    "disadvantages",
    "limitations",
    "conclusion",
    "future scope",
    "summary",
    "classification",
    "architecture",
    "components",
    "types",
    "characteristics",
    "features",
    "functions",
    "main functions",
    "working",
    "methodology",
    "system architecture",
    "parallel hardware",
    "parallel software",
    "distributed memory",
    "interconnection networks",
    "vector processors",
    "graphics processing units",
    "mimd systems",
    "simd systems"
]


def normalize_for_comparison(text):
    """
    Normalize text for heading comparison.
    """

    text = text.lower()

    text = re.sub(
        r"^\s*\d+(?:\.\d+)*[\s\.\-:]*",
        "",
        text
    )

    text = re.sub(
        r"[^a-z0-9\s]",
        " ",
        text
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def is_heading(text):
    """
    Try to identify document headings.
    """

    if not text:
        return False

    text = text.strip()

    if len(text) > 160:
        return False

    normalized = normalize_for_comparison(text)

    if not normalized:
        return False

    # Numbered headings such as:
    # 2.6 Main gateway functions
    # 3 Scope
    if re.match(
        r"^\s*\d+(?:\.\d+)*[\s\.\-:]",
        text
    ):
        return True

    # Exact keyword headings
    for keyword in HEADING_KEYWORDS:

        if normalized == keyword:
            return True

        if normalized.startswith(keyword + " "):
            return True

    # Short all-uppercase headings
    if len(text.split()) <= 10 and text.isupper():
        return True

    return False


# ============================================================
# CHUNKING
# ============================================================

def split_words(text, chunk_size=300, overlap=60):
    """
    Split text into overlapping word chunks.
    """

    words = text.split()

    if not words:
        return []

    chunks = []

    start = 0

    while start < len(words):

        end = min(
            start + chunk_size,
            len(words)
        )

        chunk = " ".join(words[start:end])

        if chunk.strip():
            chunks.append(chunk.strip())

        if end >= len(words):
            break

        start = max(
            end - overlap,
            start + 1
        )

    return chunks


def create_chunks(pages):
    """
    Create heading-aware chunks.
    """

    chunks = []

    current_heading = ""
    current_text = []
    current_page = 0

    def save_current():

        nonlocal current_text

        if not current_text:
            return

        text = "\n".join(current_text)

        text = clean_text(text)

        if not text:
            current_text = []
            return

        # Keep heading together with its content
        if current_heading:

            final_text = (
                current_heading
                + "\n"
                + text
            )

        else:

            final_text = text

        # Small sections can stay together
        word_count = len(final_text.split())

        if word_count <= 450:

            chunks.append({
                "text": final_text,
                "page": current_page,
                "heading": current_heading
            })

        else:

            smaller_chunks = split_words(
                final_text,
                chunk_size=300,
                overlap=60
            )

            for smaller in smaller_chunks:

                chunks.append({
                    "text": smaller,
                    "page": current_page,
                    "heading": current_heading
                })

        current_text = []

    # --------------------------------------------------------
    # Process pages
    # --------------------------------------------------------

    for page_data in pages:

        page_number = page_data["page"]

        lines = page_data["text"].splitlines()

        for line in lines:

            line = clean_text(line)

            if not line:
                continue

            if is_heading(line):

                save_current()

                current_heading = line
                current_page = page_number

            else:

                if current_page == 0:
                    current_page = page_number

                current_text.append(line)

    save_current()

    # --------------------------------------------------------
    # Fallback if no chunks
    # --------------------------------------------------------

    if not chunks:

        complete_text = "\n".join(
            page["text"]
            for page in pages
        )

        for part in split_words(
            complete_text,
            chunk_size=300,
            overlap=60
        ):

            chunks.append({
                "text": part,
                "page": 0,
                "heading": ""
            })

    return chunks


# ============================================================
# RESET DOCUMENT STORAGE
# ============================================================

def reset_document_storage():
    """
    Remove the old document and create a fresh collection.
    """

    global collection

    print("Resetting document storage...")

    # --------------------------------------------------------
    # Delete uploaded files
    # --------------------------------------------------------

    if os.path.exists(UPLOAD_FOLDER):

        for filename in os.listdir(UPLOAD_FOLDER):

            file_path = os.path.join(
                UPLOAD_FOLDER,
                filename
            )

            try:

                if os.path.isfile(file_path):
                    os.remove(file_path)

                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)

            except Exception as e:
                print(
                    "Could not delete:",
                    file_path,
                    e
                )

    # --------------------------------------------------------
    # Delete old Chroma collection
    # --------------------------------------------------------

    try:

        chroma_client.delete_collection(
            name="current_document"
        )

        print("Old Chroma collection deleted.")

    except Exception as e:

        print(
            "No old Chroma collection to delete:",
            e
        )

    # Important:
    # Remove old Python reference too.

    collection = None

    # --------------------------------------------------------
    # Create fresh collection
    # --------------------------------------------------------

    collection = chroma_client.create_collection(
        name="current_document",
        metadata={
            "hnsw:space": "cosine"
        },
        embedding_function=get_embedding_function()
    )

    print("Fresh Chroma collection created.")

    return collection


# ============================================================
# IMPORTANT WORDS
# ============================================================

STOPWORDS = {
    "a",
    "an",
    "the",
    "is",
    "are",
    "was",
    "were",
    "what",
    "which",
    "who",
    "when",
    "where",
    "why",
    "how",
    "does",
    "do",
    "did",
    "can",
    "could",
    "would",
    "should",
    "will",
    "about",
    "for",
    "from",
    "with",
    "and",
    "or",
    "of",
    "to",
    "in",
    "on",
    "by",
    "as",
    "at",
    "it",
    "this",
    "that",
    "these",
    "those",
    "be",
    "been",
    "being",
    "main"
}


def important_words(question):
    """
    Extract meaningful words from a question.
    """

    words = re.findall(
        r"[a-zA-Z0-9]+",
        question.lower()
    )

    return [
        word
        for word in words
        if word not in STOPWORDS
        and len(word) > 1
    ]


# ============================================================
# KEYWORD SCORE
# ============================================================

def keyword_score(question, text):
    """
    Calculate keyword overlap score.
    """

    q_words = important_words(question)

    if not q_words:
        return 0.0

    text_lower = text.lower()

    matched = 0

    for word in q_words:

        if word in text_lower:
            matched += 1

    return matched / len(q_words)


# ============================================================
# PHRASE SCORE
# ============================================================

def phrase_score(question, text):
    """
    Check exact phrase overlap.
    """

    q = clean_text(question).lower()
    t = text.lower()

    if not q:
        return 0.0

    if q in t:
        return 1.0

    words = important_words(question)

    if len(words) < 2:
        return 0.0

    phrases = []

    for i in range(len(words) - 1):

        phrases.append(
            words[i] + " " + words[i + 1]
        )

    matched = sum(
        1
        for phrase in phrases
        if phrase in t
    )

    return matched / len(phrases)


# ============================================================
# COVERAGE SCORE
# ============================================================

def coverage_score(question, text):
    """
    Measure how many important question words
    appear in the retrieved text.
    """

    words = important_words(question)

    if not words:
        return 0.0

    text_lower = text.lower()

    matched = sum(
        1
        for word in words
        if word in text_lower
    )

    return matched / len(words)


# ============================================================
# HEADING SCORE
# ============================================================

def heading_score(question, heading):
    """
    Give higher score when the question terms
    match the section heading.
    """

    if not heading:
        return 0.0

    return keyword_score(
        question,
        heading
    )


# ============================================================
# RETRIEVAL
# ============================================================

def retrieve_chunks(question):
    """
    Hybrid retrieval:
    semantic similarity + keyword + phrase + coverage + heading.
    """

    current_collection = get_collection()

    if current_collection.count() == 0:
        return []

    # --------------------------------------------------------
    # Query variations
    # --------------------------------------------------------

    q_variations = []

    original_question = clean_text(question)

    if original_question:
        q_variations.append(
            original_question
        )

    words = important_words(question)

    if words:

        cleaned_question = " ".join(words)

        if cleaned_question not in q_variations:
            q_variations.append(
                cleaned_question
            )

    # --------------------------------------------------------
    # Semantic search
    # --------------------------------------------------------

    semantic_candidates = {}

    for query in q_variations:

        try:

            results = current_collection.query(
                query_texts=[query],
                n_results=min(
                    MAX_RESULTS,
                    current_collection.count()
                ),
                include=[
                    "documents",
                    "metadatas",
                    "distances"
                ]
            )

        except Exception as e:

            print(
                "Chroma query error:",
                e
            )

            continue

        documents = results.get(
            "documents",
            [[]]
        )[0]

        metadatas = results.get(
            "metadatas",
            [[]]
        )[0]

        distances = results.get(
            "distances",
            [[]]
        )[0]

        for i, document in enumerate(documents):

            metadata = (
                metadatas[i]
                if i < len(metadatas)
                else {}
            )

            distance = (
                distances[i]
                if i < len(distances)
                else 1.0
            )

            key = document

            if key not in semantic_candidates:

                semantic_candidates[key] = {
                    "text": document,
                    "metadata": metadata,
                    "distance": distance
                }

            else:

                semantic_candidates[key]["distance"] = min(
                    semantic_candidates[key]["distance"],
                    distance
                )

    # --------------------------------------------------------
    # Lexical scan
    #
    # This is useful for exact headings such as:
    # "Main gateway functions"
    # --------------------------------------------------------

    try:

        all_data = current_collection.get(
            include=[
                "documents",
                "metadatas"
            ]
        )

        all_documents = all_data.get(
            "documents",
            []
        )

        all_metadatas = all_data.get(
            "metadatas",
            []
        )

        # Avoid processing an unnecessarily huge number
        # of documents.

        for i, document in enumerate(
            all_documents
        ):

            metadata = (
                all_metadatas[i]
                if i < len(all_metadatas)
                else {}
            )

            if document not in semantic_candidates:

                ks = keyword_score(
                    question,
                    document
                )

                ps = phrase_score(
                    question,
                    document
                )

                if ks > 0 or ps > 0:

                    semantic_candidates[document] = {
                        "text": document,
                        "metadata": metadata,
                        "distance": 1.0
                    }

    except Exception as e:

        print(
            "Lexical scan error:",
            e
        )

    # --------------------------------------------------------
    # Score candidates
    # --------------------------------------------------------

    scored = []

    for item in semantic_candidates.values():

        text = item["text"]

        metadata = item.get(
            "metadata",
            {}
        )

        distance = item.get(
            "distance",
            1.0
        )

        # Convert cosine distance into similarity-like score

        semantic_score = max(
            0.0,
            min(
                1.0,
                1.0 - distance
            )
        )

        ks = keyword_score(
            question,
            text
        )

        ps = phrase_score(
            question,
            text
        )

        cs = coverage_score(
            question,
            text
        )

        hs = heading_score(
            question,
            metadata.get(
                "heading",
                ""
            )
        )

        final_score = (
            semantic_score * 0.45
            + ks * 0.20
            + ps * 0.15
            + cs * 0.10
            + hs * 0.10
        )

        scored.append({
            "text": text,
            "metadata": metadata,
            "score": final_score,
            "semantic_score": semantic_score,
            "keyword_score": ks,
            "phrase_score": ps,
            "coverage_score": cs,
            "heading_score": hs
        })

    # --------------------------------------------------------
    # Sort
    # --------------------------------------------------------

    scored.sort(
        key=lambda x: x["score"],
        reverse=True
    )

    return scored[:MAX_RESULTS]


# ============================================================
# EXACT SECTION PRIORITY
# ============================================================

def select_context(question, retrieved):
    """
    Select the best chunks for the LLM.
    """

    if not retrieved:
        return []

    q_normalized = normalize_for_comparison(
        question
    )

    # --------------------------------------------------------
    # First check for exact heading matches
    # --------------------------------------------------------

    exact_heading_matches = []

    question_words = set(
        important_words(question)
    )

    for item in retrieved:

        heading = item["metadata"].get(
            "heading",
            ""
        )

        if not heading:
            continue

        heading_words = set(
            important_words(heading)
        )

        if not heading_words:
            continue

        overlap = len(
            question_words
            & heading_words
        )

        # Strong heading match
        if overlap >= 1:

            exact_heading_matches.append(
                (
                    overlap,
                    item
                )
            )

    exact_heading_matches.sort(
        key=lambda x: (
            x[0],
            x[1]["score"]
        ),
        reverse=True
    )

    selected = []

    # --------------------------------------------------------
    # Add best heading match first
    # --------------------------------------------------------

    for _, item in exact_heading_matches:

        if item not in selected:

            selected.append(item)

        if len(selected) >= MAX_CONTEXT_CHUNKS:
            break

    # --------------------------------------------------------
    # Add remaining high scoring chunks
    # --------------------------------------------------------

    for item in retrieved:

        if item not in selected:

            selected.append(item)

        if len(selected) >= MAX_CONTEXT_CHUNKS:
            break

    return selected


# ============================================================
# BUILD PROMPT
# ============================================================

def build_prompt(question, context_chunks):
    """
    Build a strict document-only prompt.
    """

    context_parts = []

    for index, item in enumerate(
        context_chunks,
        start=1
    ):

        metadata = item.get(
            "metadata",
            {}
        )

        heading = metadata.get(
            "heading",
            ""
        )

        page = metadata.get(
            "page",
            0
        )

        text = item.get(
            "text",
            ""
        )

        if heading:

            source_header = (
                f"Section: {heading}"
            )

        else:

            source_header = "Section: Unknown"

        if page:

            source_header += (
                f" | Page: {page}"
            )

        context_parts.append(
            f"[Document section {index}]\n"
            f"{source_header}\n"
            f"{text}"
        )

    context = "\n\n".join(
        context_parts
    )

    prompt = f"""
You are answering a question about an uploaded document.

IMPORTANT RULES:

1. Use ONLY the information provided in the document context below.
2. Do NOT use outside knowledge.
3. Do NOT guess.
4. If the answer is not supported by the document context, respond exactly:
I couldn't find this information in the currently uploaded document.
5. Keep the answer simple and easy to understand.
6. If the document gives numbered points, preserve those points.
7. Do not add information that is not present in the document.
8. Answer the question directly.

DOCUMENT CONTEXT:
{context}

QUESTION:
{question}

ANSWER:
"""

    return prompt


# ============================================================
# OLLAMA STREAM
# ============================================================

def ollama_stream(prompt):
    """
    Stream answer from Ollama Cloud.
    """

    if not OLLAMA_API_KEY:

        yield {
            "type": "error",
            "message": "OLLAMA_API_KEY is not configured."
        }

        return

    headers = {
        "Authorization": (
            f"Bearer {OLLAMA_API_KEY}"
        ),
        "Content-Type": "application/json"
    }

    payload = {
        "model": OLLAMA_MODEL,
        "prompt": prompt,
        "stream": True,
        "temperature": 0.0
    }

    try:

        response = requests.post(
            OLLAMA_URL,
            headers=headers,
            json=payload,
            stream=True,
            timeout=180
        )

        response.raise_for_status()

        for line in response.iter_lines(
            decode_unicode=True
        ):

            if not line:
                continue

            try:

                data = json.loads(line)

            except json.JSONDecodeError:

                continue

            if "response" in data:

                yield {
                    "type": "token",
                    "text": data["response"]
                }

            if data.get("done"):

                yield {
                    "type": "done"
                }

    except requests.exceptions.Timeout:

        yield {
            "type": "error",
            "message": "The AI request timed out. Please try again."
        }

    except requests.exceptions.RequestException as e:

        print(
            "Ollama request error:",
            e
        )

        yield {
            "type": "error",
            "message": "Could not connect to Ollama Cloud."
        }

    except Exception as e:

        print(
            "Ollama error:",
            e
        )

        yield {
            "type": "error",
            "message": "An unexpected error occurred."
        }


# ============================================================
# HOME PAGE
# ============================================================

@app.route("/")
def home():

    return render_template(
        "landing.html"
    )


# ============================================================
# DASHBOARD
# ============================================================

@app.route("/dashboard")
def dashboard():

    return render_template(
        "index.html"
    )


# ============================================================
# UPLOAD
# ============================================================

@app.route(
    "/upload",
    methods=["POST"]
)
def upload_file():

    global collection

    try:

        # ----------------------------------------------------
        # Check file
        # ----------------------------------------------------

        if "file" not in request.files:

            return jsonify({
                "success": False,
                "error": "No file uploaded."
            }), 400

        file = request.files["file"]

        if not file.filename:

            return jsonify({
                "success": False,
                "error": "No file selected."
            }), 400

        if not allowed_file(
            file.filename
        ):

            return jsonify({
                "success": False,
                "error": "Only PDF and DOCX files are allowed."
            }), 400

        # ----------------------------------------------------
        # Check size before processing
        # ----------------------------------------------------

        file.seek(
            0,
            os.SEEK_END
        )

        file_size = file.tell()

        file.seek(0)

        if file_size > MAX_FILE_SIZE:

            return jsonify({
                "success": False,
                "error": "File size must be 10 MB or less."
            }), 400

        if file_size == 0:

            return jsonify({
                "success": False,
                "error": "The uploaded file is empty."
            }), 400

        # ----------------------------------------------------
        # Reset old document
        # ----------------------------------------------------

        reset_document_storage()

        # ----------------------------------------------------
        # Save new file
        # ----------------------------------------------------

        filename = os.path.basename(
            file.filename
        )

        file_path = os.path.join(
            UPLOAD_FOLDER,
            filename
        )

        file.save(file_path)

        print(
            "Uploaded file:",
            filename
        )

        # ----------------------------------------------------
        # Extract text
        # ----------------------------------------------------

        pages = extract_document(
            file_path
        )

        if not pages:

            return jsonify({
                "success": False,
                "error": "Could not extract text from the document."
            }), 400

        # ----------------------------------------------------
        # Create chunks
        # ----------------------------------------------------

        chunks = create_chunks(
            pages
        )

        if not chunks:

            return jsonify({
                "success": False,
                "error": "No readable text was found in the document."
            }), 400

        print(
            "Created chunks:",
            len(chunks)
        )

        # ----------------------------------------------------
        # Get current collection
        # ----------------------------------------------------

        current_collection = get_collection()

        # ----------------------------------------------------
        # Add chunks in batches
        #
        # This helps reduce peak memory usage on Render Free.
        # ----------------------------------------------------

        batch_size = 16

        for start in range(
            0,
            len(chunks),
            batch_size
        ):

            batch = chunks[
                start:start + batch_size
            ]

            ids = [
                f"chunk_{i}"
                for i in range(
                    start,
                    start + len(batch)
                )
            ]

            documents = [
                item["text"]
                for item in batch
            ]

            metadatas = []

            for item in batch:

                metadatas.append({
                    "page": int(
                        item.get(
                            "page",
                            0
                        )
                    ),
                    "heading": item.get(
                        "heading",
                        ""
                    )
                })

            current_collection.add(
                ids=ids,
                documents=documents,
                metadatas=metadatas
            )

            print(
                f"Indexed {min(start + batch_size, len(chunks))}"
                f"/{len(chunks)} chunks"
            )

        print(
            "Document indexing completed."
        )

        return jsonify({
            "success": True,
            "filename": filename,
            "chunks": len(chunks),
            "message": "Document uploaded successfully."
        })

    except Exception as e:

        print(
            "Upload error:",
            repr(e)
        )

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


# ============================================================
# QUERY
# ============================================================

@app.route(
    "/query",
    methods=["POST"]
)
def query():

    try:

        data = request.get_json(
            silent=True
        )

        if not data:

            return jsonify({
                "success": False,
                "error": "Invalid request."
            }), 400

        question = clean_text(
            data.get(
                "question",
                ""
            )
        )

        if not question:

            return jsonify({
                "success": False,
                "error": "Please enter a question."
            }), 400

        # ----------------------------------------------------
        # Check document
        # ----------------------------------------------------

        current_collection = get_collection()

        if current_collection.count() == 0:

            return Response(
                json.dumps({
                    "type": "done",
                    "answer": NO_ANSWER
                }) + "\n",
                mimetype="application/x-ndjson"
            )

        # ----------------------------------------------------
        # Retrieve
        # ----------------------------------------------------

        retrieved = retrieve_chunks(
            question
        )

        if not retrieved:

            return Response(
                json.dumps({
                    "type": "done",
                    "answer": NO_ANSWER
                }) + "\n",
                mimetype="application/x-ndjson"
            )

        # ----------------------------------------------------
        # Select context
        # ----------------------------------------------------

        context_chunks = select_context(
            question,
            retrieved
        )

        if not context_chunks:

            return Response(
                json.dumps({
                    "type": "done",
                    "answer": NO_ANSWER
                }) + "\n",
                mimetype="application/x-ndjson"
            )

        # ----------------------------------------------------
        # Minimum relevance check
        # ----------------------------------------------------

        best_score = max(
            item["score"]
            for item in context_chunks
        )

        best_keyword = max(
            item["keyword_score"]
            for item in context_chunks
        )

        best_phrase = max(
            item["phrase_score"]
            for item in context_chunks
        )

        # If retrieval has almost no connection
        # with the question, do not ask the LLM
        # to guess.

        if (
            best_score < 0.20
            and best_keyword == 0
            and best_phrase == 0
        ):

            return Response(
                json.dumps({
                    "type": "done",
                    "answer": NO_ANSWER
                }) + "\n",
                mimetype="application/x-ndjson"
            )

        # ----------------------------------------------------
        # Build prompt
        # ----------------------------------------------------

        prompt = build_prompt(
            question,
            context_chunks
        )

        # ----------------------------------------------------
        # Stream response
        # ----------------------------------------------------

        def generate():

            full_answer = ""

            for item in ollama_stream(
                prompt
            ):

                if item["type"] == "token":

                    full_answer += item["text"]

                    yield (
                        json.dumps(item)
                        + "\n"
                    )

                elif item["type"] == "error":

                    yield (
                        json.dumps(item)
                        + "\n"
                    )

                    return

                elif item["type"] == "done":

                    # ----------------------------------------
                    # Safety check
                    # ----------------------------------------

                    cleaned_answer = (
                        full_answer.strip()
                    )

                    # If the model produced an empty answer
                    if not cleaned_answer:

                        yield (
                            json.dumps({
                                "type": "token",
                                "text": NO_ANSWER
                            })
                            + "\n"
                        )

                    # ----------------------------------------
                    # Send source information
                    # ----------------------------------------

                    sources = []

                    for item2 in context_chunks:

                        metadata = item2.get(
                            "metadata",
                            {}
                        )

                        source = {
                            "page": metadata.get(
                                "page",
                                0
                            ),
                            "heading": metadata.get(
                                "heading",
                                ""
                            )
                        }

                        if source not in sources:
                            sources.append(
                                source
                            )

                    yield (
                        json.dumps({
                            "type": "sources",
                            "sources": sources
                        })
                        + "\n"
                    )

                    yield (
                        json.dumps({
                            "type": "done"
                        })
                        + "\n"
                    )

        return Response(
            stream_with_context(
                generate()
            ),
            mimetype="application/x-ndjson"
        )

    except Exception as e:

        print(
            "Query error:",
            repr(e)
        )

        def error_stream():

            yield (
                json.dumps({
                    "type": "error",
                    "message": "An error occurred while processing your question."
                })
                + "\n"
            )

        return Response(
            stream_with_context(
                error_stream()
            ),
            mimetype="application/x-ndjson"
        )


# ============================================================
# HEALTH CHECK
# ============================================================

@app.route("/health")
def health():

    try:

        current_collection = get_collection()

        count = current_collection.count()

        return jsonify({
            "status": "ok",
            "document_loaded": count > 0,
            "chunks": count
        })

    except Exception as e:

        return jsonify({
            "status": "error",
            "error": str(e)
        }), 500


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    port = int(
        os.environ.get(
            "PORT",
            5000
        )
    )

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )