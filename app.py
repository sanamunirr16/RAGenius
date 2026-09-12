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

UPLOAD_FOLDER = "uploads"
CHROMA_FOLDER = "chroma_db"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CHROMA_FOLDER, exist_ok=True)


# ============================================================
# CONFIGURATION
# ============================================================

MAX_FILE_SIZE = 10 * 1024 * 1024

MAX_RESULTS = 30
MAX_CONTEXT_CHUNKS = 4

ALLOWED_EXTENSIONS = {
    "pdf",
    "docx"
}

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

OLLAMA_API_KEY = os.getenv(
    "OLLAMA_API_KEY"
)


# ============================================================
# CHROMA DATABASE
# ============================================================

chroma_client = chromadb.PersistentClient(
    path=CHROMA_FOLDER
)

# Lightweight ONNX-based embedding function.
# This avoids loading PyTorch + SentenceTransformer.
embedding_function = DefaultEmbeddingFunction()


def get_collection():
    """
    Get the existing current_document collection.

    The collection is NOT deleted when Flask starts.
    """

    try:
        return chroma_client.get_collection(
            name="current_document",
            embedding_function=embedding_function
        )

    except Exception:

        return chroma_client.create_collection(
            name="current_document",
            metadata={
                "hnsw:space": "cosine"
            },
            embedding_function=embedding_function
        )


collection = get_collection()


# ============================================================
# BASIC HELPERS
# ============================================================

def allowed_file(filename):

    if not filename:
        return False

    if "." not in filename:
        return False

    extension = filename.rsplit(".", 1)[1].lower()

    return extension in ALLOWED_EXTENSIONS


def clean_text(text):

    if not text:
        return ""

    text = text.replace("\x00", " ")

    text = re.sub(
        r"[ \t]+",
        " ",
        text
    )

    text = re.sub(
        r"\n{3,}",
        "\n\n",
        text
    )

    return text.strip()


# ============================================================
# PDF EXTRACTION
# ============================================================

def extract_pdf(file_path):

    pages = []

    try:

        pdf = pymupdf.open(file_path)

        for page_number, page in enumerate(
            pdf,
            start=1
        ):

            text = page.get_text("text")

            text = clean_text(text)

            if text:

                pages.append({
                    "text": text,
                    "page": page_number
                })

        pdf.close()

    except Exception as e:

        print(
            "PDF extraction error:",
            e
        )

        return []

    return pages


# ============================================================
# DOCX EXTRACTION
# ============================================================

def extract_docx(file_path):

    pages = []

    try:

        document = Document(file_path)

        content = []

        # Paragraphs
        for paragraph in document.paragraphs:

            text = clean_text(
                paragraph.text
            )

            if text:
                content.append(text)

        # Tables
        for table in document.tables:

            for row in table.rows:

                cells = []

                for cell in row.cells:

                    cell_text = clean_text(
                        cell.text
                    )

                    if cell_text:
                        cells.append(cell_text)

                if cells:

                    content.append(
                        " | ".join(cells)
                    )

        full_text = "\n".join(content)

        if full_text.strip():

            pages.append({
                "text": full_text,
                "page": 0
            })

    except Exception as e:

        print(
            "DOCX extraction error:",
            e
        )

        return []

    return pages


# ============================================================
# DOCUMENT EXTRACTION
# ============================================================

def extract_document(file_path):

    extension = file_path.rsplit(
        ".",
        1
    )[1].lower()

    if extension == "pdf":

        return extract_pdf(
            file_path
        )

    if extension == "docx":

        return extract_docx(
            file_path
        )

    return []


# ============================================================
# HEADING DETECTION
# ============================================================

HEADING_KEYWORDS = {

    "introduction",
    "definition",
    "overview",
    "importance",
    "scope",
    "applications",
    "advantages",
    "disadvantages",
    "limitations",
    "conclusion",
    "architecture",
    "methodology",
    "components",
    "features",
    "functions",
    "types",
    "classification",
    "characteristics",
    "working",
    "working principle",
    "objectives",
    "benefits",
    "challenges",
    "future scope",
    "future work",

    "parallel hardware",
    "parallel software",
    "distributed memory",
    "interconnection networks",
    "vector processors",
    "graphics processing units",
    "mimd systems",
    "simd systems",

    "m2m gateway",
    "gateway functions",

    "software defined networking",
    "software-defined networking"
}


def normalize_for_comparison(text):

    text = text.lower()

    text = text.replace(
        "software-defined",
        "software defined"
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


def is_heading(line):

    line = line.strip()

    if not line:
        return False

    if len(line) > 120:
        return False

    normalized = normalize_for_comparison(
        line
    )

    # Numbered headings
    # Examples:
    # 1. Introduction
    # 2.3 Architecture
    # 2.6 Main gateway functions

    if re.match(
        r"^\d+(?:\.\d+)*\s+[A-Za-z]",
        line
    ):

        return True

    # Examples:
    # 1) Introduction
    # 2- Architecture

    if re.match(
        r"^\d+[\)\-]\s+[A-Za-z]",
        line
    ):

        return True

    # Known headings

    for keyword in HEADING_KEYWORDS:

        if normalized == keyword:
            return True

        if normalized.startswith(
            keyword + " "
        ):
            return True

    # Short ALL CAPS heading

    if (
        len(line.split()) <= 12
        and line.upper() == line
        and any(
            ch.isalpha()
            for ch in line
        )
    ):

        return True

    return False


# ============================================================
# WORD CHUNKING
# ============================================================

def split_words(
    text,
    chunk_size=300,
    overlap=60
):

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

        chunk = " ".join(
            words[start:end]
        ).strip()

        if chunk:

            chunks.append(
                chunk
            )

        if end >= len(words):
            break

        start = max(
            end - overlap,
            start + 1
        )

    return chunks


# ============================================================
# CREATE SMART CHUNKS
# ============================================================

def create_chunks(pages):

    all_chunks = []

    for page_data in pages:

        page_text = page_data["text"]
        page_number = page_data["page"]

        lines = page_text.splitlines()

        current_heading = ""
        current_section = []

        sections = []

        # Detect sections

        for line in lines:

            line = line.strip()

            if not line:
                continue

            if is_heading(line):

                if current_section:

                    sections.append({
                        "heading": current_heading,
                        "text": " ".join(
                            current_section
                        ),
                        "page": page_number
                    })

                current_heading = line

                current_section = []

            else:

                current_section.append(
                    line
                )

        # Last section

        if current_section:

            sections.append({
                "heading": current_heading,
                "text": " ".join(
                    current_section
                ),
                "page": page_number
            })

        # No headings

        if not sections:

            sections = [{
                "heading": "",
                "text": page_text,
                "page": page_number
            }]

        # Create chunks

        for section in sections:

            heading = section["heading"]

            text = clean_text(
                section["text"]
            )

            if not text:
                continue

            words = text.split()

            # Keep normal sections together

            if len(words) <= 450:

                chunk_text = text

                if heading:

                    chunk_text = (
                        heading
                        + "\n"
                        + chunk_text
                    )

                all_chunks.append({

                    "text": chunk_text,

                    "heading": heading,

                    "page": section["page"]

                })

            # Split large sections

            else:

                split_chunks = split_words(
                    text,
                    chunk_size=300,
                    overlap=60
                )

                for chunk in split_chunks:

                    chunk_text = chunk

                    if heading:

                        chunk_text = (
                            heading
                            + "\n"
                            + chunk
                        )

                    all_chunks.append({

                        "text": chunk_text,

                        "heading": heading,

                        "page": section["page"]

                    })

    return all_chunks


# ============================================================
# RESET DOCUMENT STORAGE
# ============================================================

def reset_document_storage():

    global collection

    # Delete uploaded files

    try:

        for filename in os.listdir(
            UPLOAD_FOLDER
        ):

            file_path = os.path.join(
                UPLOAD_FOLDER,
                filename
            )

            if os.path.isfile(file_path):

                os.remove(file_path)

            elif os.path.isdir(file_path):

                shutil.rmtree(file_path)

    except Exception as e:

        print(
            "Upload folder reset error:",
            e
        )

    # Delete old collection

    try:

        chroma_client.delete_collection(
            name="current_document"
        )

    except Exception:

        pass

    # Create fresh collection

    collection = chroma_client.create_collection(

        name="current_document",

        metadata={
            "hnsw:space": "cosine"
        },

        embedding_function=embedding_function
    )


# ============================================================
# STOP WORDS
# ============================================================

STOP_WORDS = {

    "what",
    "is",
    "are",
    "was",
    "were",
    "the",
    "a",
    "an",
    "of",
    "for",
    "to",
    "in",
    "on",
    "and",
    "or",
    "with",
    "by",
    "from",
    "how",
    "why",
    "where",
    "when",
    "which",
    "who",
    "does",
    "do",
    "can",
    "could",
    "would",
    "should",
    "explain",
    "give",
    "tell",
    "about",
    "define",
    "definition",
    "describe",
    "main",
    "following",
    "following:",
    "please"
}


# ============================================================
# IMPORTANT WORDS
# ============================================================

def important_words(question):

    words = re.findall(
        r"[a-zA-Z0-9]+",
        question.lower()
    )

    result = []

    for word in words:

        if len(word) <= 1:
            continue

        if word in STOP_WORDS:
            continue

        result.append(word)

    return result


# ============================================================
# KEYWORD SCORE
# ============================================================

def keyword_score(
    question,
    text
):

    query_words = important_words(
        question
    )

    if not query_words:
        return 0.0

    normalized_text = normalize_for_comparison(
        text
    )

    matches = 0

    for word in query_words:

        if word in normalized_text:

            matches += 1

    return matches / len(
        query_words
    )


# ============================================================
# PHRASE SCORE
# ============================================================

def phrase_score(
    question,
    text
):

    q = normalize_for_comparison(
        question
    )

    t = normalize_for_comparison(
        text
    )

    if not q:
        return 0.0

    if q in t:
        return 1.0

    words = important_words(
        question
    )

    if len(words) >= 2:

        phrase = " ".join(
            words
        )

        if phrase in t:
            return 0.9

    return 0.0


# ============================================================
# COVERAGE SCORE
# ============================================================

def coverage_score(
    question,
    text
):

    words = important_words(
        question
    )

    if not words:
        return 0.0

    text_normalized = normalize_for_comparison(
        text
    )

    found = 0

    for word in words:

        if word in text_normalized:
            found += 1

    return found / len(words)


# ============================================================
# QUERY SECTION DETECTION
# ============================================================

def extract_possible_section(question):

    q = normalize_for_comparison(
        question
    )

    possible_sections = [

        "main gateway functions",

        "gateway functions",

        "m2m gateway",

        "parallel hardware",

        "parallel software",

        "distributed memory",

        "interconnection networks",

        "graphics processing units",

        "vector processors",

        "mimd systems",

        "simd systems",

        "software defined networking",

        "software defined network"
    ]

    for section in possible_sections:

        if section in q:

            return section

    return None


# ============================================================
# HEADING MATCH
# ============================================================

def heading_matches_query(
    question,
    heading
):

    if not heading:
        return False

    q = normalize_for_comparison(
        question
    )

    h = normalize_for_comparison(
        heading
    )

    if not h:
        return False

    if h in q:
        return True

    heading_words = [

        word

        for word in h.split()

        if word not in STOP_WORDS

    ]

    if not heading_words:
        return False

    matched = 0

    for word in heading_words:

        if word in q:
            matched += 1

    return matched >= max(
        1,
        len(heading_words) // 2
    )


# ============================================================
# RETRIEVE CHUNKS
# ============================================================

def retrieve_chunks(question):

    if collection.count() == 0:

        print(
            "No chunks in Chroma."
        )

        return []

    print(
        "Retrieving for:",
        question
    )

    # Query variations

    query_variations = [
        question
    ]

    words = important_words(
        question
    )

    if words:

        important_query = " ".join(
            words
        )

        if (
            important_query.lower()
            != question.lower()
        ):

            query_variations.append(
                important_query
            )

    results_map = {}

    total_chunks = collection.count()

    search_count = min(
        MAX_RESULTS,
        total_chunks
    )

    # --------------------------------------------------------
    # Semantic search using Chroma's built-in embeddings
    # --------------------------------------------------------

    for query_text in query_variations:

        try:

            results = collection.query(

                query_texts=[
                    query_text
                ],

                n_results=search_count,

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

        for index, document in enumerate(
            documents
        ):

            metadata = (

                metadatas[index]

                if index < len(metadatas)

                else {}

            )

            distance = (

                distances[index]

                if index < len(distances)

                else 1.0

            )

            similarity = max(

                0.0,

                min(

                    1.0,

                    1.0 - float(distance)

                )

            )

            if document not in results_map:

                results_map[document] = {

                    "text": document,

                    "heading": metadata.get(
                        "heading",
                        ""
                    ),

                    "page": metadata.get(
                        "page",
                        0
                    ),

                    "filename": metadata.get(
                        "filename",
                        ""
                    ),

                    "semantic": similarity

                }

            else:

                results_map[
                    document
                ]["semantic"] = max(

                    results_map[
                        document
                    ]["semantic"],

                    similarity

                )

    # --------------------------------------------------------
    # EXACT KEYWORD SCAN
    # --------------------------------------------------------

    try:

        all_data = collection.get(

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

        for index, document in enumerate(
            all_documents
        ):

            metadata = (

                all_metadatas[index]

                if index < len(all_metadatas)

                else {}

            )

            if document not in results_map:

                k_score = keyword_score(
                    question,
                    document
                )

                p_score = phrase_score(
                    question,
                    document
                )

                if (
                    k_score > 0
                    or p_score > 0
                ):

                    results_map[document] = {

                        "text": document,

                        "heading": metadata.get(
                            "heading",
                            ""
                        ),

                        "page": metadata.get(
                            "page",
                            0
                        ),

                        "filename": metadata.get(
                            "filename",
                            ""
                        ),

                        "semantic": 0.0

                    }

    except Exception as e:

        print(
            "Lexical scan error:",
            e
        )

    # --------------------------------------------------------
    # SCORE CANDIDATES
    # --------------------------------------------------------

    candidates = []

    for item in results_map.values():

        text = item["text"]

        heading = item["heading"]

        semantic = item.get(
            "semantic",
            0.0
        )

        k_score = keyword_score(
            question,
            text
        )

        p_score = phrase_score(
            question,
            text
        )

        c_score = coverage_score(
            question,
            text
        )

        heading_score = 0.0

        if heading_matches_query(
            question,
            heading
        ):

            heading_score = 1.0

        score = (

            semantic * 0.35

            + k_score * 0.30

            + p_score * 0.20

            + c_score * 0.10

            + heading_score * 0.05

        )

        item["keyword"] = k_score

        item["phrase"] = p_score

        item["coverage"] = c_score

        item["heading_score"] = heading_score

        item["score"] = score

        candidates.append(item)

    if not candidates:

        print(
            "No candidates found."
        )

        return []

    # --------------------------------------------------------
    # EXACT SECTION PRIORITY
    # --------------------------------------------------------

    exact_section = extract_possible_section(
        question
    )

    if exact_section:

        exact_candidates = []

        for item in candidates:

            heading = normalize_for_comparison(
                item.get(
                    "heading",
                    ""
                )
            )

            text = normalize_for_comparison(
                item.get(
                    "text",
                    ""
                )
            )

            if (
                exact_section in heading
                or exact_section in text
            ):

                exact_candidates.append(
                    item
                )

        if exact_candidates:

            candidates = exact_candidates

    # --------------------------------------------------------
    # SORT
    # --------------------------------------------------------

    candidates.sort(

        key=lambda item: (

            item["keyword"],

            item["phrase"],

            item["coverage"],

            item["heading_score"],

            item["score"],

            item["semantic"]

        ),

        reverse=True
    )

    # --------------------------------------------------------
    # DEBUG
    # --------------------------------------------------------

    print(
        "Top retrieved chunks:"
    )

    for item in candidates[:5]:

        print(

            "Score:",
            round(
                item["score"],
                3
            ),

            "| Keyword:",
            round(
                item["keyword"],
                3
            ),

            "| Heading:",
            item["heading"]

        )

    return candidates[:MAX_RESULTS]


# ============================================================
# SELECT FINAL CONTEXT
# ============================================================

def select_context(
    question,
    candidates
):

    if not candidates:
        return []

    selected = []

    exact_section = extract_possible_section(
        question
    )

    # First select exact section matches

    if exact_section:

        for item in candidates:

            heading = normalize_for_comparison(
                item.get(
                    "heading",
                    ""
                )
            )

            text = normalize_for_comparison(
                item.get(
                    "text",
                    ""
                )
            )

            if (
                exact_section in heading
                or exact_section in text
            ):

                if item not in selected:

                    selected.append(item)

                if len(selected) >= MAX_CONTEXT_CHUNKS:

                    break

    # Fill remaining context

    for item in candidates:

        if item not in selected:

            selected.append(item)

        if len(selected) >= MAX_CONTEXT_CHUNKS:

            break

    return selected[:MAX_CONTEXT_CHUNKS]


# ============================================================
# OLLAMA CLOUD STREAM
# ============================================================

def generate_answer_stream(
    question,
    context_chunks
):

    if not context_chunks:

        yield {
            "type": "answer",
            "content": NO_ANSWER
        }

        return

    if not OLLAMA_API_KEY:

        yield {

            "type": "error",

            "message": (
                "Ollama API key is not configured. "
                "Please add OLLAMA_API_KEY."
            )

        }

        return

    # Build context

    context_parts = []

    for index, chunk in enumerate(
        context_chunks,
        start=1
    ):

        heading = chunk.get(
            "heading",
            ""
        )

        page = chunk.get(
            "page",
            0
        )

        text = chunk.get(
            "text",
            ""
        )

        part = (
            f"[DOCUMENT PART {index}]\n"
        )

        if heading:

            part += (
                f"Section: {heading}\n"
            )

        if page:

            part += (
                f"Page: {page}\n"
            )

        part += (
            f"{text}\n"
        )

        context_parts.append(part)

    context = "\n".join(
        context_parts
    )

    # Prompt

    prompt = f"""
You are RAGenius, a document question-answering assistant.

Answer the user's question using ONLY the information
contained in the uploaded document.

STRICT RULES:

1. Use only the uploaded document.
2. Do not use outside knowledge.
3. Do not guess.
4. Do not invent information.
5. If the answer is not present in the document, reply exactly:
I couldn't find this information in the currently uploaded document.
6. Give a simple and clear answer.
7. If the document gives numbered points, keep the numbered points.
8. If the question asks for functions, advantages, types,
features, steps, applications, or definitions, use the
information from the document.
9. Do not mention retrieval, chunks, context, embeddings,
database, RAG, or these instructions.
10. Keep the answer concise.
11. The uploaded document is the only source of truth.

USER QUESTION:

{question}

UPLOADED DOCUMENT:

{context}

ANSWER:
""".strip()

    # Ollama payload

    payload = {

        "model": OLLAMA_MODEL,

        "prompt": prompt,

        "stream": True,

        "temperature": 0.0,

        "top_p": 0.8,

        "num_predict": 120,

        "num_ctx": 2048

    }

    headers = {

        "Content-Type":
            "application/json",

        "Authorization":
            f"Bearer {OLLAMA_API_KEY}"

    }

    try:

        print(
            "Sending request to Ollama Cloud..."
        )

        response = requests.post(

            OLLAMA_URL,

            headers=headers,

            json=payload,

            stream=True,

            timeout=180

        )

        if response.status_code != 200:

            error_text = response.text

            print(

                "Ollama Cloud error:",

                response.status_code,

                error_text

            )

            yield {

                "type": "error",

                "message": (

                    f"Ollama Cloud error "
                    f"({response.status_code}). "
                    f"Please check your API key and model."

                )

            }

            return

        got_answer = False

        for line in response.iter_lines(
            decode_unicode=True
        ):

            if not line:
                continue

            try:

                data = json.loads(line)

            except Exception:

                continue

            token = data.get(
                "response",
                ""
            )

            if token:

                got_answer = True

                yield {

                    "type": "token",

                    "content": token

                }

            if data.get(
                "done",
                False
            ):

                break

        if not got_answer:

            yield {

                "type": "answer",

                "content": NO_ANSWER

            }

    except requests.exceptions.Timeout:

        yield {

            "type": "error",

            "message":
                "The AI service took too long to respond."

        }

    except requests.exceptions.ConnectionError:

        yield {

            "type": "error",

            "message":
                "Could not connect to Ollama Cloud."

        }

    except Exception as e:

        print(
            "Ollama generation error:",
            e
        )

        yield {

            "type": "error",

            "message":
                "An error occurred while generating the answer."

        }


# ============================================================
# LANDING PAGE
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

    try:

        if "file" not in request.files:

            return jsonify({

                "success": False,

                "error":
                    "No file selected."

            }), 400

        file = request.files["file"]

        if not file.filename:

            return jsonify({

                "success": False,

                "error":
                    "No file selected."

            }), 400

        if not allowed_file(
            file.filename
        ):

            return jsonify({

                "success": False,

                "error":
                    "Only PDF and DOCX files are supported."

            }), 400

        file.seek(
            0,
            os.SEEK_END
        )

        file_size = file.tell()

        file.seek(0)

        if file_size > MAX_FILE_SIZE:

            return jsonify({

                "success": False,

                "error":
                    "File size must be less than 10 MB."

            }), 400

        # New document replaces old one

        reset_document_storage()

        safe_filename = os.path.basename(
            file.filename
        )

        file_path = os.path.join(
            UPLOAD_FOLDER,
            safe_filename
        )

        file.save(file_path)

        print(
            "Uploaded:",
            safe_filename
        )

        # Extract

        pages = extract_document(
            file_path
        )

        if not pages:

            return jsonify({

                "success": False,

                "error":
                    "Could not extract text from the uploaded document."

            }), 400

        # Create chunks

        chunks = create_chunks(
            pages
        )

        if not chunks:

            return jsonify({

                "success": False,

                "error":
                    "No readable text was found in the document."

            }), 400

        print(
            "Created chunks:",
            len(chunks)
        )

        # ----------------------------------------------------
        # Store documents in Chroma.
        #
        # Chroma now creates embeddings automatically.
        # No SentenceTransformer/PyTorch is loaded.
        # ----------------------------------------------------

        texts = [

            chunk["text"]

            for chunk in chunks

        ]

        ids = []

        metadatas = []

        for index, chunk in enumerate(
            chunks
        ):

            ids.append(
                f"chunk_{index}"
            )

            metadatas.append({

                "heading":
                    chunk.get(
                        "heading",
                        ""
                    ),

                "page":
                    int(
                        chunk.get(
                            "page",
                            0
                        )
                    ),

                "filename":
                    safe_filename

            })

        collection.add(

            ids=ids,

            documents=texts,

            metadatas=metadatas

        )

        print(
            "Stored chunks:",
            collection.count()
        )

        return jsonify({

            "success": True,

            "message":
                "Document uploaded successfully.",

            "filename":
                safe_filename,

            "chunks":
                len(chunks)

        })

    except Exception as e:

        print(
            "UPLOAD ERROR:",
            e
        )

        return jsonify({

            "success": False,

            "error":
                str(e)

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

                "answer":
                    NO_ANSWER

            }), 400

        question = str(
            data.get(
                "question",
                ""
            )
        ).strip()

        if not question:

            return jsonify({

                "answer":
                    "Please enter a question."

            }), 400

        document_count = collection.count()

        print(
            "Current document chunks:",
            document_count
        )

        if document_count == 0:

            return jsonify({

                "answer":
                    NO_ANSWER

            })

        # Retrieve

        candidates = retrieve_chunks(
            question
        )

        if not candidates:

            return jsonify({

                "answer":
                    NO_ANSWER

            })

        # Select context

        context_chunks = select_context(
            question,
            candidates
        )

        if not context_chunks:

            return jsonify({

                "answer":
                    NO_ANSWER

            })

        print(
            "Context chunks selected:",
            len(context_chunks)
        )

        # Source information

        filename = ""

        pages = []

        for chunk in context_chunks:

            if not filename:

                filename = chunk.get(
                    "filename",
                    ""
                )

            page = chunk.get(
                "page",
                0
            )

            if page:

                pages.append(page)

        pages = sorted(
            list(set(pages))
        )

        # Streaming response

        def generate():

            source = {

                "filename":
                    filename,

                "pages":
                    pages

            }

            yield (

                json.dumps({

                    "type":
                        "source",

                    "source":
                        source

                })

                + "\n"

            )

            for item in generate_answer_stream(

                question,

                context_chunks

            ):

                yield (

                    json.dumps(item)

                    + "\n"

                )

        return Response(

            stream_with_context(
                generate()
            ),

            mimetype=
                "application/x-ndjson"

        )

    except Exception as e:

        print(
            "QUERY ERROR:",
            e
        )

        return jsonify({

            "answer":
                "An error occurred while processing your question."

        }), 500


# ============================================================
# HEALTH CHECK
# ============================================================

@app.route("/health")
def health():

    return jsonify({

        "status":
            "ok",

        "document_chunks":
            collection.count(),

        "ollama_model":
            OLLAMA_MODEL,

        "ollama_key_configured":
            bool(OLLAMA_API_KEY)

    })


# ============================================================
# RUN APPLICATION
# ============================================================

if __name__ == "__main__":

    app.run(

        host="0.0.0.0",

        port=int(
            os.environ.get(
                "PORT",
                5000
            )
        ),

        debug=False

    )