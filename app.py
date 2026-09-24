from flask import Flask, render_template, request, jsonify
import os
import re
import shutil
import fitz
from docx import Document
from difflib import SequenceMatcher


# ============================================================
# APP
# ============================================================

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")

MAX_FILE_SIZE = 10 * 1024 * 1024  # 10 MB

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_SIZE

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ============================================================
# CONSTANTS
# ============================================================

ALLOWED_EXTENSIONS = {"pdf", "docx"}

NOT_FOUND = (
    "I couldn't find this information in the currently uploaded document."
)


# ============================================================
# ONE ACTIVE DOCUMENT ONLY
# ============================================================

current_document = {
    "filename": None,
    "filetype": None,
    "pages": [],
    "blocks": [],
    "lines": [],
    "sections": [],
    "items": []
}


# ============================================================
# TEXT CLEANING
# ============================================================

def clean_text(text):
    if text is None:
        return ""

    text = str(text)

    text = text.replace("\xa0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\ufeff", "")

    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n[ \t]+", "\n", text)

    return text.strip()


def normalize_text(text):
    text = clean_text(text).lower()

    text = text.replace("&", " and ")

    text = re.sub(r"[^a-z0-9\s]", " ", text)

    text = re.sub(r"\s+", " ", text)

    return text.strip()


# ============================================================
# QUESTION PROCESSING
# ============================================================

QUESTION_PREFIXES = [
    "what is",
    "what are",
    "what was",
    "what were",
    "who is",
    "who are",
    "who was",
    "where is",
    "where are",
    "when is",
    "when was",
    "why is",
    "why are",
    "why was",
    "why were",
    "how is",
    "how are",
    "how does",
    "how do",
    "how can",
    "define",
    "definition of",
    "meaning of",
    "explain",
    "explain about",
    "describe",
    "tell me about",
    "tell me",
    "tell about",
    "give information about",
    "give information on",
    "give details about",
    "give details on",
    "write about",
    "discuss",
    "elaborate on",
    "elaborate about",
    "can you explain",
    "can you tell me about",
    "briefly explain",
    "briefly describe"
]


def question_topic(question):
    """
    Converts:

        What is machine learning?

    into:

        machine learning
    """

    text = clean_text(question)

    text = text.rstrip(" ?!:")

    lower = text.lower()

    for prefix in sorted(
        QUESTION_PREFIXES,
        key=len,
        reverse=True
    ):
        if lower.startswith(prefix + " "):
            text = text[len(prefix):].strip()
            break

        if lower == prefix:
            text = ""
            break

    text = text.rstrip(" ?!:")

    return clean_text(text)


# ============================================================
# PREFIX / BULLET CLEANING
# ============================================================

def remove_number_prefix(text):
    """
    Removes:

    1.
    1.1
    1.2.3
    2)
    """

    return re.sub(
        r"^\s*\d+(?:\.\d+)*[\.\)]?\s*",
        "",
        clean_text(text)
    ).strip()


def remove_letter_prefix(text):
    """
    Removes:

    a.
    b)
    A.
    (a)
    """

    text = clean_text(text)

    text = re.sub(
        r"^\s*\([A-Za-z0-9]+\)\s*",
        "",
        text
    )

    text = re.sub(
        r"^\s*[A-Za-z][\.\)]\s*",
        "",
        text
    )

    return text.strip()


def remove_bullet_marker(text):
    text = clean_text(text)

    return re.sub(
        r"^\s*(?:[-•●▪◦‣*]|[oO])\s+",
        "",
        text
    ).strip()


def clean_label(text):
    text = clean_text(text)

    text = remove_bullet_marker(text)
    text = remove_number_prefix(text)
    text = remove_letter_prefix(text)

    return text.strip(" :.-")


# ============================================================
# BULLET DETECTION
# ============================================================

def is_bullet(text):
    if not text:
        return False

    text = clean_text(text)

    return bool(
        re.match(
            r"^(?:[-•●▪◦‣*]|[oO])\s+[A-Za-z0-9]",
            text
        )
    )


# ============================================================
# HEADING DETECTION
# ============================================================

def numbered_heading_info(text):
    """
    Returns:

        (True, level)

    for:

        1. Introduction       -> level 1
        1.1 Definition        -> level 2
        1.1.1 Example         -> level 3
        a. Types              -> level 2

    Otherwise:

        (False, None)
    """

    text = clean_text(text)

    # 1. Introduction
    match = re.match(
        r"^\s*(\d+)\.\s+.+",
        text
    )

    if match:
        return True, 1

    # 1.1 Definition
    match = re.match(
        r"^\s*(\d+(?:\.\d+)+)\s+.+",
        text
    )

    if match:
        parts = match.group(1).split(".")
        return True, len(parts)

    # a. Definition
    if re.match(
        r"^\s*[A-Za-z][\.\)]\s+.+",
        text
    ):
        return True, 2

    # (a) Definition
    if re.match(
        r"^\s*\([A-Za-z0-9]+\)\s+.+",
        text
    ):
        return True, 2

    return False, None


def is_all_caps_heading(text):
    text = clean_text(text)

    if not text:
        return False

    letters = re.sub(
        r"[^A-Za-z]",
        "",
        text
    )

    if len(letters) < 4:
        return False

    upper = sum(
        1 for char in letters
        if char.isupper()
    )

    ratio = upper / len(letters)

    if ratio >= 0.80 and len(text.split()) <= 16:
        return True

    return False


def is_colon_heading(text):
    """
    Detects short headings ending with ':'.

    Example:

        Applications:
        Advantages:
        Classification:
    """

    text = clean_text(text)

    if not text.endswith(":"):
        return False

    left = text[:-1].strip()

    if not left:
        return False

    if len(left.split()) > 10:
        return False

    return True


def looks_like_title(text):
    """
    Generic fallback for headings such as:

        Introduction
        Applications
        Machine Learning Algorithms
        Data Preprocessing
        Improved Surface Finish and Accuracy
    """

    text = clean_text(text)

    if not text:
        return False

    if len(text) > 100:
        return False

    if is_bullet(text):
        return False

    if text.endswith("."):
        return False

    words = text.split()

    if len(words) > 12:
        return False

    # Avoid ordinary sentence-like paragraphs
    sentence_words = {
        "the",
        "this",
        "these",
        "those",
        "is",
        "are",
        "was",
        "were",
        "can",
        "could",
        "may",
        "might",
        "will",
        "would",
        "has",
        "have",
        "had",
        "using",
        "used"
    }

    lower_words = [
        re.sub(
            r"[^a-z]",
            "",
            word.lower()
        )
        for word in words
    ]

    sentence_word_count = sum(
        1
        for word in lower_words
        if word in sentence_words
    )

    if len(words) >= 6 and sentence_word_count >= 2:
        return False

    # Title case signal
    title_case_count = 0

    for word in words:

        word_clean = re.sub(
            r"[^A-Za-z]",
            "",
            word
        )

        if not word_clean:
            continue

        if word_clean[0].isupper():
            title_case_count += 1

    title_ratio = (
        title_case_count /
        max(len(words), 1)
    )

    return title_ratio >= 0.60


def generic_heading_detection(
    text,
    style_name="",
    bold_ratio=0.0
):
    """
    Generic heading detector.

    It does NOT know anything about NTM,
    Machine Learning, Big Data, DBMS, etc.
    """

    text = clean_text(text)

    if not text:
        return False, None

    style_name = style_name.lower()

    # --------------------------------------------------------
    # Word heading styles
    # --------------------------------------------------------

    match = re.search(
        r"heading\s*(\d+)",
        style_name
    )

    if match:
        return True, int(match.group(1))

    if "title" in style_name:
        return True, 1

    if "subtitle" in style_name:
        return True, 2

    # --------------------------------------------------------
    # Numbered headings
    # --------------------------------------------------------

    numbered, level = numbered_heading_info(text)

    if numbered:
        return True, level

    # --------------------------------------------------------
    # Strong bold short paragraph
    # --------------------------------------------------------

    if bold_ratio >= 0.75:

        if len(text.split()) <= 15:
            return True, 2

    # --------------------------------------------------------
    # ALL CAPS
    # --------------------------------------------------------

    if is_all_caps_heading(text):
        return True, 1

    # --------------------------------------------------------
    # Short colon heading
    # --------------------------------------------------------

    if is_colon_heading(text):
        return True, 2

    # --------------------------------------------------------
    # Title-like heading
    # --------------------------------------------------------

    if looks_like_title(text):
        return True, 2

    return False, None


# ============================================================
# INLINE "LABEL: ANSWER"
# ============================================================

def parse_inline_item(text):
    """
    Detects:

        Definition: Machine learning is...

        1. Definition: Machine learning is...

        - Advantages: ...
    """

    original = clean_text(text)

    if not original:
        return None

    candidate = remove_bullet_marker(original)
    candidate = remove_number_prefix(candidate)
    candidate = remove_letter_prefix(candidate)

    if ":" not in candidate:
        return None

    left, right = candidate.split(
        ":",
        1
    )

    left = clean_text(left)
    right = clean_text(right)

    if not left or not right:
        return None

    # A label should normally be short
    if len(left.split()) > 12:
        return None

    return {
        "label": clean_label(left),
        "answer": right
    }


# ============================================================
# DOCX PARAGRAPH INFORMATION
# ============================================================

def get_docx_style(paragraph):
    try:
        return paragraph.style.name or ""
    except Exception:
        return ""


def get_bold_ratio(paragraph):
    runs = [
        run
        for run in paragraph.runs
        if clean_text(run.text)
    ]

    if not runs:
        return 0.0

    bold_runs = sum(
        1
        for run in runs
        if run.bold is True
    )

    return bold_runs / len(runs)


def get_docx_heading_info(paragraph):
    text = clean_text(paragraph.text)

    style_name = get_docx_style(
        paragraph
    )

    bold_ratio = get_bold_ratio(
        paragraph
    )

    return generic_heading_detection(
        text,
        style_name,
        bold_ratio
    )


# ============================================================
# RESET
# ============================================================

def reset_document():
    global current_document

    current_document = {
        "filename": None,
        "filetype": None,
        "pages": [],
        "blocks": [],
        "lines": [],
        "sections": [],
        "items": []
    }


def clear_upload_folder():

    os.makedirs(
        UPLOAD_FOLDER,
        exist_ok=True
    )

    for name in os.listdir(
        UPLOAD_FOLDER
    ):

        path = os.path.join(
            UPLOAD_FOLDER,
            name
        )

        try:

            if os.path.isfile(path):
                os.remove(path)

            elif os.path.isdir(path):
                shutil.rmtree(path)

        except Exception as error:
            print(
                "Could not remove:",
                path,
                error
            )


# ============================================================
# PDF EXTRACTION
# ============================================================

def extract_pdf(path):

    pages = []
    blocks = []
    lines = []

    document = fitz.open(path)

    try:

        for page_number, page in enumerate(
            document,
            start=1
        ):

            page_blocks = page.get_text(
                "blocks"
            )

            page_lines = []

            for block in page_blocks:

                if len(block) < 5:
                    continue

                text = clean_text(
                    block[4]
                )

                block_type = (
                    block[6]
                    if len(block) > 6
                    else 0
                )

                # Ignore image blocks
                if block_type != 0:
                    continue

                if not text:
                    continue

                page_lines.append(
                    text
                )

                blocks.append({
                    "page": page_number,
                    "text": text
                })

                # ------------------------------------------------
                # IMPORTANT:
                # Preserve ALL lines from PDF blocks.
                # Do not truncate them.
                # ------------------------------------------------

                for one_line in text.split(
                    "\n"
                ):

                    one_line = clean_text(
                        one_line
                    )

                    if one_line:
                        lines.append({
                            "page": page_number,
                            "text": one_line,
                            "heading": False,
                            "heading_level": None
                        })

            pages.append({
                "page": page_number,
                "text": "\n".join(
                    page_lines
                )
            })

    finally:
        document.close()

    # ------------------------------------------------------------
    # Generic PDF heading detection
    # ------------------------------------------------------------

    for line in lines:

        text = line["text"]

        heading, level = (
            generic_heading_detection(
                text
            )
        )

        line["heading"] = heading
        line["heading_level"] = level

    return (
        pages,
        blocks,
        lines
    )


# ============================================================
# DOCX EXTRACTION
# ============================================================

def extract_docx(path):

    pages = []
    blocks = []
    lines = []

    document = Document(path)

    # --------------------------------------------------------
    # Paragraphs
    # --------------------------------------------------------

    for paragraph in document.paragraphs:

        text = clean_text(
            paragraph.text
        )

        if not text:
            continue

        heading, level = (
            get_docx_heading_info(
                paragraph
            )
        )

        blocks.append({
            "page": 1,
            "text": text,
            "heading": heading,
            "heading_level": level
        })

        lines.append({
            "page": 1,
            "text": text,
            "heading": heading,
            "heading_level": level
        })

    # --------------------------------------------------------
    # Tables
    # --------------------------------------------------------

    for table in document.tables:

        for row in table.rows:

            cells = []

            for cell in row.cells:

                cell_text = clean_text(
                    cell.text
                )

                if cell_text:
                    cells.append(
                        cell_text
                    )

            if not cells:
                continue

            row_text = " | ".join(
                cells
            )

            blocks.append({
                "page": 1,
                "text": row_text,
                "heading": False,
                "heading_level": None
            })

            lines.append({
                "page": 1,
                "text": row_text,
                "heading": False,
                "heading_level": None
            })

    pages.append({
        "page": 1,
        "text": "\n".join(
            line["text"]
            for line in lines
        )
    })

    return (
        pages,
        blocks,
        lines
    )


# ============================================================
# BUILD HIERARCHICAL SECTIONS
# ============================================================

def build_sections(lines):

    sections = []

    stack = []

    current_section = None

    for index, line in enumerate(
        lines
    ):

        text = clean_text(
            line.get("text", "")
        )

        if not text:
            continue

        is_heading = line.get(
            "heading",
            False
        )

        level = line.get(
            "heading_level"
        )

        # ----------------------------------------------------
        # HEADING
        # ----------------------------------------------------

        if is_heading:

            if level is None:
                level = 2

            # Save old section
            if current_section:

                answer = clean_text(
                    "\n".join(
                        current_section[
                            "content"
                        ]
                    )
                )

                current_section[
                    "answer"
                ] = answer

                if answer:
                    sections.append(
                        current_section
                    )

            # Remove deeper hierarchy
            while stack and stack[-1][
                "level"
            ] >= level:
                stack.pop()

            parent = (
                stack[-1]
                if stack
                else None
            )

            section = {
                "label": clean_label(
                    text
                ),
                "raw_label": text,
                "level": level,
                "page": line.get(
                    "page",
                    1
                ),
                "line_index": index,
                "content": [],
                "answer": "",
                "parent": (
                    parent["label"]
                    if parent
                    else ""
                )
            }

            stack.append(
                section
            )

            current_section = section

            continue

        # ----------------------------------------------------
        # NORMAL CONTENT
        # ----------------------------------------------------

        if current_section is None:

            # Document text before first heading
            current_section = {
                "label": "",
                "raw_label": "",
                "level": 99,
                "page": line.get(
                    "page",
                    1
                ),
                "line_index": index,
                "content": [],
                "answer": "",
                "parent": ""
            }

        current_section[
            "content"
        ].append(text)

    # Save final section
    if current_section:

        answer = clean_text(
            "\n".join(
                current_section[
                    "content"
                ]
            )
        )

        current_section[
            "answer"
        ] = answer

        if answer:
            sections.append(
                current_section
            )

    return sections


# ============================================================
# BUILD SEARCH ITEMS
# ============================================================

def build_items(lines):

    items = []

    current_heading = ""
    current_level = 99
    current_parent = ""

    pending = None

    def save_pending():

        nonlocal pending

        if not pending:
            return

        answer = clean_text(
            " ".join(
                pending["parts"]
            )
        )

        if answer:

            items.append({
                "label": pending[
                    "label"
                ],
                "answer": answer,
                "page": pending[
                    "page"
                ],
                "line_index": pending[
                    "line_index"
                ],
                "heading": current_heading,
                "heading_level": current_level,
                "parent": current_parent
            })

        pending = None

    for index, line in enumerate(
        lines
    ):

        text = clean_text(
            line.get("text", "")
        )

        if not text:
            continue

        # ----------------------------------------------------
        # NEW HEADING
        # ----------------------------------------------------

        if line.get(
            "heading",
            False
        ):

            save_pending()

            current_heading = clean_label(
                text
            )

            current_level = line.get(
                "heading_level",
                2
            )

            # Find nearest previous parent heading
            current_parent = ""

            for previous in reversed(
                items
            ):

                previous_heading = clean_text(
                    previous.get(
                        "heading",
                        ""
                    )
                )

                previous_level = previous.get(
                    "heading_level",
                    99
                )

                if (
                    previous_heading
                    and previous_level
                    < current_level
                ):

                    current_parent = (
                        previous_heading
                    )

                    break

            continue

        # ----------------------------------------------------
        # BULLET
        # ----------------------------------------------------

        if is_bullet(text):

            save_pending()

            cleaned = remove_bullet_marker(
                text
            )

            inline = parse_inline_item(
                cleaned
            )

            if inline:

                pending = {
                    "label": inline[
                        "label"
                    ],
                    "parts": [
                        inline[
                            "answer"
                        ]
                    ],
                    "page": line.get(
                        "page",
                        1
                    ),
                    "line_index": index
                }

            else:

                pending = {
                    "label": "",
                    "parts": [cleaned],
                    "page": line.get(
                        "page",
                        1
                    ),
                    "line_index": index
                }

            continue

        # ----------------------------------------------------
        # INLINE LABEL: ANSWER
        # ----------------------------------------------------

        inline = parse_inline_item(
            text
        )

        if inline:

            save_pending()

            pending = {
                "label": inline[
                    "label"
                ],
                "parts": [
                    inline[
                        "answer"
                    ]
                ],
                "page": line.get(
                    "page",
                    1
                ),
                "line_index": index
            }

            continue

        # ----------------------------------------------------
        # CONTINUATION
        # ----------------------------------------------------

        if pending:

            pending[
                "parts"
            ].append(text)

        else:

            # Normal paragraph
            items.append({
                "label": "",
                "answer": text,
                "page": line.get(
                    "page",
                    1
                ),
                "line_index": index,
                "heading": current_heading,
                "heading_level": current_level,
                "parent": current_parent
            })

    save_pending()

    return [
        item
        for item in items
        if clean_text(
            item.get(
                "answer",
                ""
            )
        )
    ]


# ============================================================
# ADD SECTION ITEMS
# ============================================================

def add_section_items(
    items,
    sections
):
    """
    Creates searchable entries for complete
    document sections.

    This is generic and works for any subject.
    """

    result = list(items)

    for section in sections:

        label = clean_text(
            section.get(
                "label",
                ""
            )
        )

        answer = clean_text(
            section.get(
                "answer",
                ""
            )
        )

        if not label or not answer:
            continue

        normalized = normalize_text(
            label
        )

        # Find same label + same parent
        found = False

        for item in result:

            item_label = normalize_text(
                item.get(
                    "label",
                    ""
                )
            )

            item_parent = normalize_text(
                item.get(
                    "parent",
                    ""
                )
            )

            section_parent = normalize_text(
                section.get(
                    "parent",
                    ""
                )
            )

            if (
                item_label == normalized
                and item_parent
                == section_parent
            ):

                # Keep the more complete answer
                if len(answer) > len(
                    clean_text(
                        item.get(
                            "answer",
                            ""
                        )
                    )
                ):

                    item["answer"] = answer

                found = True
                break

        if not found:

            result.append({
                "label": label,
                "answer": answer,
                "page": section.get(
                    "page",
                    1
                ),
                "line_index": section.get(
                    "line_index",
                    0
                ),
                "heading": label,
                "heading_level": section.get(
                    "level",
                    2
                ),
                "parent": section.get(
                    "parent",
                    ""
                )
            })

    return result


# ============================================================
# TOKENIZATION
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
    "to",
    "for",
    "in",
    "on",
    "at",
    "and",
    "or",
    "with",
    "about",
    "does",
    "do",
    "can",
    "could",
    "would",
    "should",
    "please",
    "me",
    "tell",
    "give",
    "explain",
    "define",
    "describe",
    "write",
    "discuss"
}


def tokens(text):

    words = normalize_text(
        text
    ).split()

    return {
        word
        for word in words
        if (
            word not in STOP_WORDS
            and len(word) > 1
        )
    }


def similarity(a, b):

    a = normalize_text(a)
    b = normalize_text(b)

    if not a or not b:
        return 0.0

    return SequenceMatcher(
        None,
        a,
        b
    ).ratio()


def overlap_score(a, b):

    a_tokens = tokens(a)
    b_tokens = tokens(b)

    if not a_tokens:
        return 0.0

    return len(
        a_tokens & b_tokens
    ) / len(a_tokens)


def exact_match(a, b):

    return (
        normalize_text(a)
        ==
        normalize_text(b)
    )


# ============================================================
# FIND EXACT HEADING
# ============================================================

def find_exact_heading(question):

    topic = question_topic(
        question
    )

    if not topic:
        return None

    matches = []

    for item in current_document[
        "items"
    ]:

        label = clean_text(
            item.get(
                "label",
                ""
            )
        )

        answer = clean_text(
            item.get(
                "answer",
                ""
            )
        )

        if not label or not answer:
            continue

        if exact_match(
            topic,
            label
        ):

            matches.append(
                item
            )

    if not matches:
        return None

    # --------------------------------------------------------
    # If duplicate labels exist, prefer the one with:
    #
    # 1. More specific hierarchy
    # 2. Longer complete answer
    #
    # This is still entirely document-based.
    # --------------------------------------------------------

    matches.sort(
        key=lambda item: (
            item.get(
                "heading_level",
                99
            ),
            -len(
                clean_text(
                    item.get(
                        "answer",
                        ""
                    )
                )
            )
        )
    )

    best = matches[0]

    return {
        "type": "exact_heading",
        "score": 1000,
        "page": best.get(
            "page",
            1
        ),
        "label": best.get(
            "label",
            ""
        ),
        "answer": best.get(
            "answer",
            ""
        ),
        "line_index": best.get(
            "line_index",
            0
        )
    }


# ============================================================
# FIND STRONG CONTAINMENT
# ============================================================

def find_heading_contains(question):

    topic = question_topic(
        question
    )

    if not topic:
        return None

    topic_norm = normalize_text(
        topic
    )

    candidates = []

    for item in current_document[
        "items"
    ]:

        label = clean_text(
            item.get(
                "label",
                ""
            )
        )

        answer = clean_text(
            item.get(
                "answer",
                ""
            )
        )

        if not label or not answer:
            continue

        label_norm = normalize_text(
            label
        )

        if (
            topic_norm in label_norm
            or label_norm in topic_norm
        ):

            candidates.append(
                item
            )

    if not candidates:
        return None

    # Require meaningful containment
    candidates.sort(
        key=lambda item: (
            len(
                normalize_text(
                    item.get(
                        "label",
                        ""
                    )
                )
            ),
            -len(
                item.get(
                    "answer",
                    ""
                )
            )
        )
    )

    best = candidates[0]

    return {
        "type": "heading_contains",
        "score": 900,
        "page": best.get(
            "page",
            1
        ),
        "label": best.get(
            "label",
            ""
        ),
        "answer": best.get(
            "answer",
            ""
        ),
        "line_index": best.get(
            "line_index",
            0
        )
    }


# ============================================================
# FUZZY HEADING
# ============================================================

def find_similar_heading(question):

    topic = question_topic(
        question
    )

    if not topic:
        return None

    candidates = []

    for item in current_document[
        "items"
    ]:

        label = clean_text(
            item.get(
                "label",
                ""
            )
        )

        answer = clean_text(
            item.get(
                "answer",
                ""
            )
        )

        if not label or not answer:
            continue

        sim = similarity(
            topic,
            label
        )

        overlap = overlap_score(
            topic,
            label
        )

        if (
            sim >= 0.90
            and overlap >= 0.65
        ):

            score = (
                sim * 100
                +
                overlap * 50
            )

            candidates.append({
                "score": score,
                "item": item
            })

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: x["score"],
        reverse=True
    )

    # --------------------------------------------------------
    # Important:
    # If two candidates are almost equally good,
    # DO NOT GUESS.
    # --------------------------------------------------------

    if len(candidates) > 1:

        difference = (
            candidates[0]["score"]
            -
            candidates[1]["score"]
        )

        if difference < 4:
            return None

    best = candidates[0]["item"]

    return {
        "type": "similar_heading",
        "score": candidates[0][
            "score"
        ],
        "page": best.get(
            "page",
            1
        ),
        "label": best.get(
            "label",
            ""
        ),
        "answer": best.get(
            "answer",
            ""
        ),
        "line_index": best.get(
            "line_index",
            0
        )
    }


# ============================================================
# LEXICAL RETRIEVAL
# ============================================================

def lexical_search(question):

    topic = question_topic(
        question
    )

    if not topic:
        return None

    query_tokens = tokens(
        topic
    )

    if not query_tokens:
        return None

    candidates = []

    for item in current_document[
        "items"
    ]:

        label = clean_text(
            item.get(
                "label",
                ""
            )
        )

        answer = clean_text(
            item.get(
                "answer",
                ""
            )
        )

        if not answer:
            continue

        label_tokens = tokens(
            label
        )

        answer_tokens = tokens(
            answer
        )

        label_overlap = 0

        answer_overlap = 0

        if query_tokens:

            label_overlap = (
                len(
                    query_tokens
                    &
                    label_tokens
                )
                /
                len(query_tokens)
            )

            answer_overlap = (
                len(
                    query_tokens
                    &
                    answer_tokens
                )
                /
                len(query_tokens)
            )

        # Label match is much stronger
        score = (
            label_overlap * 100
            +
            answer_overlap * 30
        )

        if normalize_text(
            topic
        ) in normalize_text(
            label
        ):
            score += 100

        if score > 0:

            candidates.append({
                "score": score,
                "item": item
            })

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: x["score"],
        reverse=True
    )

    best_score = candidates[0][
        "score"
    ]

    # --------------------------------------------------------
    # Strict confidence check
    # --------------------------------------------------------

    if best_score < 55:
        return None

    # Avoid random answer when candidates are too close
    if len(candidates) > 1:

        second_score = candidates[1][
            "score"
        ]

        if (
            best_score - second_score
            < 8
        ):
            return None

    best = candidates[0]["item"]

    return {
        "type": "lexical",
        "score": best_score,
        "page": best.get(
            "page",
            1
        ),
        "label": best.get(
            "label",
            ""
        ),
        "answer": best.get(
            "answer",
            ""
        ),
        "line_index": best.get(
            "line_index",
            0
        )
    }


# ============================================================
# SECTION RETRIEVAL
# ============================================================

def retrieve_matching_section(
    question
):

    topic = question_topic(
        question
    )

    if not topic:
        return None

    candidates = []

    for section in current_document[
        "sections"
    ]:

        label = clean_text(
            section.get(
                "label",
                ""
            )
        )

        answer = clean_text(
            section.get(
                "answer",
                ""
            )
        )

        if not label or not answer:
            continue

        sim = similarity(
            topic,
            label
        )

        overlap = overlap_score(
            topic,
            label
        )

        if (
            sim >= 0.90
            and overlap >= 0.65
        ):

            score = (
                sim * 80
                +
                overlap * 50
            )

            candidates.append({
                "score": score,
                "section": section
            })

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: x["score"],
        reverse=True
    )

    if len(candidates) > 1:

        if (
            candidates[0]["score"]
            -
            candidates[1]["score"]
            < 5
        ):
            return None

    best = candidates[0]["section"]

    return {
        "type": "section",
        "score": candidates[0][
            "score"
        ],
        "page": best.get(
            "page",
            1
        ),
        "label": best.get(
            "label",
            ""
        ),
        "answer": best.get(
            "answer",
            ""
        ),
        "line_index": best.get(
            "line_index",
            0
        )
    }


# ============================================================
# MAIN RETRIEVAL
# ============================================================

def retrieve_answer(question):

    if not current_document[
        "filename"
    ]:

        return {
            "type": "none",
            "score": 0,
            "page": None,
            "label": "",
            "answer": NOT_FOUND
        }

    # --------------------------------------------------------
    # 1. Exact heading
    # --------------------------------------------------------

    result = find_exact_heading(
        question
    )

    if result:
        return result

    # --------------------------------------------------------
    # 2. Strong heading containment
    # --------------------------------------------------------

    result = find_heading_contains(
        question
    )

    if result:
        return result

    # --------------------------------------------------------
    # 3. High confidence fuzzy heading
    # --------------------------------------------------------

    result = find_similar_heading(
        question
    )

    if result:
        return result

    # --------------------------------------------------------
    # 4. Strong lexical search
    # --------------------------------------------------------

    result = lexical_search(
        question
    )

    if result:
        return result

    # --------------------------------------------------------
    # 5. Section search
    # --------------------------------------------------------

    result = retrieve_matching_section(
        question
    )

    if result:
        return result

    # --------------------------------------------------------
    # 6. NOT FOUND
    # --------------------------------------------------------

    return {
        "type": "none",
        "score": 0,
        "page": None,
        "label": "",
        "answer": NOT_FOUND
    }


# ============================================================
# HOME
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

    global current_document

    file = request.files.get(
        "file"
    )

    if file is None:

        file = request.files.get(
            "document"
        )

    if not file or not file.filename:

        return jsonify({
            "success": False,
            "message": "No file selected."
        }), 400

    filename = os.path.basename(
        file.filename
    )

    extension = ""

    if "." in filename:

        extension = (
            filename
            .rsplit(
                ".",
                1
            )[1]
            .lower()
        )

    if extension not in ALLOWED_EXTENSIONS:

        return jsonify({
            "success": False,
            "message": (
                "Only PDF and DOCX files are supported."
            )
        }), 400

    # --------------------------------------------------------
    # IMPORTANT:
    # COMPLETELY REMOVE OLD DOCUMENT
    # --------------------------------------------------------

    reset_document()

    clear_upload_folder()

    # --------------------------------------------------------
    # SAFE FILE NAME
    # --------------------------------------------------------

    safe_filename = re.sub(
        r"[^A-Za-z0-9_.-]",
        "_",
        filename
    )

    filepath = os.path.join(
        UPLOAD_FOLDER,
        safe_filename
    )

    try:

        file.save(
            filepath
        )

        # ----------------------------------------------------
        # EXTRACT
        # ----------------------------------------------------

        if extension == "pdf":

            pages, blocks, lines = (
                extract_pdf(
                    filepath
                )
            )

        else:

            pages, blocks, lines = (
                extract_docx(
                    filepath
                )
            )

        # ----------------------------------------------------
        # BUILD DOCUMENT STRUCTURE
        # ----------------------------------------------------

        sections = build_sections(
            lines
        )

        items = build_items(
            lines
        )

        items = add_section_items(
            items,
            sections
        )

        # ----------------------------------------------------
        # STORE ONLY THIS DOCUMENT
        # ----------------------------------------------------

        current_document = {
            "filename": filename,
            "filetype": extension,
            "pages": pages,
            "blocks": blocks,
            "lines": lines,
            "sections": sections,
            "items": items
        }

        # ----------------------------------------------------
        # DEBUG INFORMATION
        # ----------------------------------------------------

        print("\n================================")
        print("NEW DOCUMENT LOADED")
        print("================================")
        print(
            "Filename:",
            filename
        )
        print(
            "Type:",
            extension
        )
        print(
            "Pages:",
            len(pages)
        )
        print(
            "Lines:",
            len(lines)
        )
        print(
            "Sections:",
            len(sections)
        )
        print(
            "Items:",
            len(items)
        )

        print("\nDETECTED SECTIONS:")

        for section in sections:

            print(
                f"  LEVEL {section.get('level')}: "
                f"{section.get('label')}"
            )

        print(
            "================================\n"
        )

        return jsonify({
            "success": True,
            "message": (
                "Document uploaded successfully."
            ),
            "filename": filename,
            "filetype": extension,
            "pages": len(pages),
            "sections": len(sections),
            "items": len(items)
        })

    except Exception as error:

        print(
            "DOCUMENT PROCESSING ERROR:",
            error
        )

        reset_document()

        try:

            if os.path.exists(
                filepath
            ):
                os.remove(
                    filepath
                )

        except Exception:
            pass

        return jsonify({
            "success": False,
            "message": (
                "Could not process the document."
            )
        }), 500


# ============================================================
# QUERY
# ============================================================

@app.route(
    "/query",
    methods=["POST"]
)
def query_document():

    if not current_document[
        "filename"
    ]:

        return jsonify({
            "answer": NOT_FOUND,
            "source": {
                "filename": None,
                "page": None,
                "type": "none",
                "label": ""
            }
        })

    data = request.get_json(
        silent=True
    )

    if not data:

        return jsonify({
            "answer": NOT_FOUND,
            "source": {
                "filename": current_document[
                    "filename"
                ],
                "page": None,
                "type": "none",
                "label": ""
            }
        })

    question = clean_text(
        data.get(
            "question",
            ""
        )
    )

    if not question:

        return jsonify({
            "answer": NOT_FOUND,
            "source": {
                "filename": current_document[
                    "filename"
                ],
                "page": None,
                "type": "none",
                "label": ""
            }
        })

    print(
        "\nQUESTION:",
        question
    )

    result = retrieve_answer(
        question
    )

    answer = clean_text(
        result.get(
            "answer",
            NOT_FOUND
        )
    )

    if not answer:
        answer = NOT_FOUND

    page = result.get(
        "page"
    )

    result_type = result.get(
        "type",
        "none"
    )

    label = result.get(
        "label",
        ""
    )

    print(
        "MATCH:",
        label
    )

    print(
        "TYPE:",
        result_type
    )

    print(
        "SCORE:",
        result.get(
            "score",
            0
        )
    )

    print(
        "ANSWER:",
        answer
    )

    return jsonify({
        "answer": answer,
        "source": {
            "filename": current_document[
                "filename"
            ],
            "page": page,
            "type": result_type,
            "label": label
        }
    })


# ============================================================
# CURRENT DOCUMENT STATUS
# ============================================================

@app.route(
    "/document",
    methods=["GET"]
)
def document_status():

    if not current_document[
        "filename"
    ]:

        return jsonify({
            "filename": None,
            "filetype": None,
            "pages": 0,
            "sections": 0,
            "items": 0
        })

    return jsonify({
        "filename": current_document[
            "filename"
        ],
        "filetype": current_document[
            "filetype"
        ],
        "pages": len(
            current_document[
                "pages"
            ]
        ),
        "sections": len(
            current_document[
                "sections"
            ]
        ),
        "items": len(
            current_document[
                "items"
            ]
        )
    })


# ============================================================
# CLEAR DOCUMENT
# ============================================================

@app.route(
    "/clear",
    methods=["POST"]
)
def clear_document():

    reset_document()

    clear_upload_folder()

    return jsonify({
        "success": True,
        "message": (
            "Current document removed."
        )
    })


# ============================================================
# FILE TOO LARGE
# ============================================================

@app.errorhandler(413)
def file_too_large(error):

    return jsonify({
        "success": False,
        "message": (
            "File is too large. "
            "Maximum size is 10 MB."
        )
    }), 413


# ============================================================
# RUN
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
        debug=True
    )