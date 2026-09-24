/* =========================================================
   RAGenius
   Frontend JavaScript
========================================================= */


/* =========================================================
   ELEMENTS
========================================================= */

const fileInput =
    document.getElementById("fileInput");

const questionInput =
    document.getElementById("questionInput");

const chatMessages =
    document.getElementById("chatMessages");

const chatWelcome =
    document.getElementById("chatWelcome");

const documentList =
    document.getElementById("documentList");

const emptyDocuments =
    document.getElementById("emptyDocuments");

const documentCount =
    document.getElementById("documentCount");

const footerDocumentCount =
    document.getElementById("footerDocumentCount");


/* =========================================================
   DOCUMENT STORAGE
========================================================= */

let documents = [];


/* =========================================================
   OPEN FILE PICKER
========================================================= */

function openFilePicker() {

    if (!fileInput) {
        return;
    }

    fileInput.click();
}


/* =========================================================
   FILE SELECTED
========================================================= */

if (fileInput) {

    fileInput.addEventListener(
        "change",
        function () {

            if (this.files.length > 0) {

                uploadFile(this.files[0]);

                this.value = "";

            }

        }
    );

}


/* =========================================================
   UPLOAD FILE
========================================================= */

async function uploadFile(file) {


    /* Check extension */

    const extension =
        file.name
            .split(".")
            .pop()
            .toLowerCase();


    if (
        extension !== "pdf" &&
        extension !== "docx"
    ) {

        alert(
            "Please upload only PDF or DOCX files."
        );

        return;
    }


    /* Check file size */

    if (
        file.size >
        10 * 1024 * 1024
    ) {

        alert(
            "File size should be less than 10 MB."
        );

        return;
    }


    /* Clear previous chat */

    chatMessages.innerHTML = "";

    showChatWelcome();


    /* Show temporary uploading message */

    const uploading =
        addMessage(
            "Uploading " + file.name + "...",
            "ai",
            true
        );


    const formData =
        new FormData();

    formData.append(
        "file",
        file
    );


    try {

        const response =
            await fetch(
                "/upload",
                {
                    method: "POST",
                    body: formData
                }
            );


        const data =
            await response.json();


        if (!response.ok) {

            throw new Error(
                data.error ||
                "Upload failed."
            );
        }


        /* Remove uploading message */

        if (uploading) {
            uploading.remove();
        }


        /* Add document */

        addDocument(file.name);


        /* Show successful upload */

        addMessage(
            "✓ " +
            file.name +
            " uploaded successfully. You can now ask questions about this document.",
            "system"
        );


    } catch (error) {

        console.error(
            "Upload error:",
            error
        );


        if (uploading) {
            uploading.remove();
        }


        addMessage(
            "⚠ " +
            (
                error.message ||
                "Could not upload the document."
            ),
            "system"
        );

    }

}


/* =========================================================
   ADD DOCUMENT
========================================================= */

function addDocument(filename) {


    /* Only one active document */

    documents = [filename];


    documentList.innerHTML = "";


    emptyDocuments.style.display =
        "none";


    const item =
        document.createElement("div");


    item.className =
        "document-item";


    item.innerHTML = `

        <div class="document-item-icon">
            📄
        </div>

        <div
            class="document-item-details">

            <div
                class="document-item-name"
                title="${escapeHtml(filename)}">

                ${escapeHtml(filename)}

            </div>

            <div
                class="document-item-status">

                Ready to analyze

            </div>

        </div>

    `;


    documentList.prepend(item);


    updateDocumentCount();

}


/* =========================================================
   DOCUMENT COUNT
========================================================= */

function updateDocumentCount() {

    const count =
        documents.length;


    documentCount.textContent =
        count;


    footerDocumentCount.textContent =
        count;

}


/* =========================================================
   HIDE CHAT WELCOME
========================================================= */

function hideChatWelcome() {

    if (!chatWelcome) {
        return;
    }


    chatWelcome.classList.add(
        "hidden"
    );

}


/* =========================================================
   SHOW CHAT WELCOME
========================================================= */

function showChatWelcome() {

    if (!chatWelcome) {
        return;
    }


    chatWelcome.classList.remove(
        "hidden"
    );

}


/* =========================================================
   SEND QUESTION
========================================================= */

async function sendQuestion() {


    if (!questionInput) {
        return;
    }


    const question =
        questionInput.value.trim();


    /* Empty question */

    if (!question) {
        return;
    }


    /* No document */

    if (documents.length === 0) {

        hideChatWelcome();


        addMessage(
            "Please upload a PDF or DOCX document first.",
            "system"
        );


        questionInput.value = "";


        return;
    }


    /* Hide brain/logo */

    hideChatWelcome();


    /* Add user's question */

    addMessage(
        question,
        "user"
    );


    /* Clear input */

    questionInput.value = "";


    /* Add loading box */

    const loading =
        addMessage(
            "Searching your document...",
            "ai",
            true
        );


    try {


        const response =
            await fetch(
                "/query",
                {

                    method: "POST",

                    headers: {
                        "Content-Type":
                            "application/json"
                    },

                    body:
                        JSON.stringify({
                            question:
                                question
                        })

                }
            );


        /* The backend always answers with a plain JSON body,
           even for "not found", so read it as JSON rather
           than as a token stream. */

        const data =
            await response.json();


        if (!response.ok) {

            throw new Error(
                data.answer ||
                "Server could not process the question."
            );

        }


        /* Remove loading */

        if (loading) {
            loading.remove();
        }


        /* Create AI answer box */

        const answerText =
            (data.answer || "").trim() ||
            "I couldn't find an answer in the currently uploaded document.";


        const answerMessage =
            addMessage(
                answerText,
                "ai"
            );


        if (data.source) {

            answerMessage.dataset.source =
                JSON.stringify(
                    data.source
                );

        }


        scrollChatToBottom();


    } catch (error) {


        console.error(
            "Question error:",
            error
        );


        if (loading) {
            loading.remove();
        }


        addMessage(
            "Something went wrong while contacting the RAG engine.",
            "system"
        );

    }

}


/* =========================================================
   ENTER KEY
========================================================= */

if (questionInput) {

    questionInput.addEventListener(
        "keydown",
        function (event) {

            if (
                event.key === "Enter" &&
                !event.shiftKey
            ) {

                event.preventDefault();

                sendQuestion();

            }

        }
    );

}


/* =========================================================
   ADD MESSAGE
========================================================= */

function addMessage(
    text,
    type,
    temporary = false
) {


    const message =
        document.createElement("div");


    message.classList.add(
        "message"
    );


    if (
        type === "user"
    ) {

        message.classList.add(
            "user-message"
        );

    }

    else if (
        type === "ai"
    ) {

        message.classList.add(
            "ai-message"
        );

    }

    else {

        message.classList.add(
            "system-message"
        );

    }


    if (temporary) {

        message.classList.add(
            "message-loading"
        );

    }


    message.textContent =
        text;


    chatMessages.appendChild(
        message
    );


    scrollChatToBottom();


    return message;

}


/* =========================================================
   SCROLL CHAT
========================================================= */

function scrollChatToBottom() {

    if (!chatMessages) {
        return;
    }


    chatMessages.scrollTop =
        chatMessages.scrollHeight;

}


/* =========================================================
   NEW CONVERSATION
========================================================= */

function newConversation() {


    chatMessages.innerHTML = "";


    questionInput.value = "";


    showChatWelcome();


    scrollChatToBottom();

}


/* =========================================================
   SETTINGS
========================================================= */

function showSettings() {

    hideChatWelcome();


    addMessage(
        "Settings panel will be available here.",
        "system"
    );

}


/* =========================================================
   HELP
========================================================= */

function showHelp() {

    hideChatWelcome();


    addMessage(
        "Upload a PDF or DOCX file, then ask RAGenius questions about the information inside your document.",
        "system"
    );

}


/* =========================================================
   ESCAPE HTML
========================================================= */

function escapeHtml(text) {

    const div =
        document.createElement(
            "div"
        );


    div.textContent =
        text;


    return div.innerHTML;

}


/* =========================================================
   INITIAL STATE
========================================================= */

updateDocumentCount();

showChatWelcome();