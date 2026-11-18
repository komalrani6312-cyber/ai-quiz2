"""
AI Quiz & Question Generator  (Python + Streamlit + Gemini + ReportLab)

BACKEND FLOW (kya-kya hota hai):
  1. Student sidebar mein topic, type, count, difficulty deta hai
  2. build_prompt()        -> in sab se ek instruction (prompt) banta hai
  3. generate_questions()  -> prompt Gemini ko jata hai, JSON wapas aata hai
  4. Validation            -> galat format wale questions hata diye jate hain
  5. session_state         -> questions memory mein save rehte hain
  6. Checkbox selection    -> student jo chahe wo questions chunta hai
  7. build_pdf()           -> sirf selected questions, Q1 se dobara numbering
  8. download_button       -> PDF download
"""

import json
import os
import time
from html import escape as html_escape
from io import BytesIO
from xml.sax.saxutils import escape

import streamlit as st
from dotenv import load_dotenv
from google import genai
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer
from storage import (
    DuplicateAttemptError,
    DuplicateUsernameError,
    authenticate_teacher,
    build_marksheet_xlsx,
    create_note,
    create_teacher,
    create_quiz,
    get_attempt_review,
    get_attempts,
    get_quiz,
    get_student_result,
    get_note,
    list_cheating_events,
    list_notes,
    list_quizzes,
    list_published_quizzes,
    list_shared_notes,
    record_cheating_event,
    submit_attempt,
    init_db,
)

load_dotenv()

LETTERS = ["A", "B", "C", "D"]
QUESTION_TYPES = ["MCQ", "Short Answer", "Long Answer"]
DEFAULT_MODEL = "gemini-3.5-flash"
FALLBACK_MODELS = ["gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-3.8-flash"]


def get_api_key():
    """Get the Gemini key from .env or Streamlit secrets, without exposing it in the UI."""
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass
    return os.getenv("GEMINI_API_KEY", "")


def get_config_value(name):
    try:
        value = st.secrets.get(name, "")
    except Exception:
        value = ""
    return str(value or os.getenv(name, "")).strip()


def has_placeholder_key(api_key_value):
    if not api_key_value:
        return True
    cleaned = str(api_key_value).strip()
    blacklist = [
        "your_api_key_here",
        "your_api_key",
        "api_key_here",
        "example",
        "replace_me",
        "test_key",
    ]
    lower = cleaned.lower()
    return lower in blacklist or "your_api" in lower or "api_key_here" in lower or len(cleaned) < 20


# ==========================================================================
# PART 1 : AI (Gemini) - questions generate karna
# ==========================================================================
def build_prompt(topic, count, q_type, difficulty, extra):
    """Student ki requirement se AI ke liye instruction banata hai."""
    if q_type == "MCQ":
        fmt = (
            'Each item: {"question": str, "options": [4 strings without A/B/C/D prefix], '
            '"answer": "A" or "B" or "C" or "D", "explanation": str (1 line)}'
        )
    else:
        length = "2-3 line" if q_type == "Short Answer" else "detailed 8-10 line"
        fmt = f'Each item: {{"question": str, "answer": str ({length} model answer)}}'

    return f"""You are an expert exam paper setter.
Create exactly {count} {q_type} questions on the topic: "{topic}".
Difficulty: {difficulty}.
Extra instructions from student: {extra or "None"}.
Questions must be different from each other and factually correct.
Return ONLY a JSON array (no markdown, no extra text). {fmt}"""


def generate_questions(api_key, model, topic, count, q_type, difficulty, extra):
    """Gemini ko call karta hai, JSON parse karta hai, galat questions hataata hai."""
    client = genai.Client(api_key=api_key)
    prompt = build_prompt(topic, count, q_type, difficulty, extra)

    # Retry a temporarily busy model once, then try the configured fallbacks.
    models_to_try = [model] + [m for m in FALLBACK_MODELS if m != model]
    response, last_error = None, None
    for m in models_to_try:
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=m, contents=prompt,
                    config={"response_mime_type": "application/json"},
                )
                break
            except Exception as error:
                last_error = error
                message = str(error).upper()
                is_missing_model = "404" in message or "NOT_FOUND" in message
                is_temporary = any(
                    marker in message
                    for marker in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED", "RATE LIMIT")
                )
                if is_missing_model:
                    break
                if is_temporary:
                    if attempt == 0:
                        time.sleep(1)
                        continue
                    break
                raise
        if response is not None:
            break
    if response is None:
        raise last_error

    text = response.text.strip().replace("```json", "").replace("```", "")
    data = json.loads(text)

    clean = []                                   # validation
    for q in data:
        if not isinstance(q, dict) or not q.get("question"):
            continue
        if q_type == "MCQ":
            opts = q.get("options", [])
            if len(opts) != 4 or q.get("answer") not in LETTERS:
                continue
        clean.append(q)
    return clean


# ==========================================================================
# PART 2 : PDF export (ReportLab)
# ==========================================================================
def _page_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 9)
    canvas.setFillColor(colors.grey)
    canvas.drawCentredString(A4[0] / 2, 1.1 * cm, f"Page {doc.page}")
    canvas.restoreState()


def build_pdf(title, questions, q_type, with_answers):
    """questions = sirf selected questions. Numbering 1 se dobara start hoti hai."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, title=title,
        leftMargin=2 * cm, rightMargin=2 * cm, topMargin=2 * cm, bottomMargin=2 * cm,
    )
    base = getSampleStyleSheet()
    purple = colors.HexColor("#4B3FCF")
    title_style = ParagraphStyle("T", parent=base["Title"], textColor=purple, fontSize=22)
    sub_style = ParagraphStyle("S", parent=base["Normal"], textColor=colors.grey, alignment=1)
    q_style = ParagraphStyle("Q", parent=base["Normal"], fontSize=11, leading=15, spaceBefore=10)
    o_style = ParagraphStyle("O", parent=base["Normal"], fontSize=10.5, leading=14, leftIndent=22)
    h_style = ParagraphStyle("H", parent=base["Heading1"], textColor=purple)

    story = [
        Paragraph(escape(title), title_style),
        Paragraph(f"Total Questions: {len(questions)}", sub_style),
        Spacer(1, 14),
    ]
    for n, q in enumerate(questions, start=1):
        story.append(Paragraph(f"<b>Q{n}.</b> {escape(q['question'])}", q_style))
        if q_type == "MCQ":
            for letter, opt in zip(LETTERS, q["options"]):
                story.append(Paragraph(f"({letter}) {escape(str(opt))}", o_style))
        else:
            story.append(Spacer(1, 40 if q_type == "Short Answer" else 120))

    if with_answers:
        story += [Spacer(1, 24), Paragraph("Answer Key", h_style)]
        for n, q in enumerate(questions, start=1):
            if q_type == "MCQ":
                idx = LETTERS.index(q["answer"])
                text = f"<b>Q{n}.</b> ({q['answer']}) {escape(str(q['options'][idx]))}"
                if q.get("explanation"):
                    text += f"<br/><i>{escape(q['explanation'])}</i>"
            else:
                text = f"<b>Q{n}.</b> {escape(q.get('answer', ''))}"
            story += [Paragraph(text, o_style), Spacer(1, 4)]

    doc.build(story, onFirstPage=_page_footer, onLaterPages=_page_footer)
    return buffer.getvalue()


# ==========================================================================
# PART 3 : Selection helpers (session_state mein checkbox ki state)
# ==========================================================================
def set_all(value):
    for i in range(len(st.session_state.questions)):
        st.session_state[f"teacher_sel_{i}"] = value


def pick_first(n):
    for i in range(len(st.session_state.questions)):
        st.session_state[f"teacher_sel_{i}"] = i < n


def selected_questions():
    return [
        q for i, q in enumerate(st.session_state.questions)
        if st.session_state.get(f"teacher_sel_{i}", False)
    ]


def set_access_role(role):
    st.session_state.access_role = role


# ==========================================================================
# PART 4 : Look & feel (CSS)
# ==========================================================================
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"], .stMarkdown, button, input, textarea {font-family:'Inter',sans-serif;}
[data-testid="stAppDeployButton"], .stDeployButton {display:none;}
section.main > div.block-container{padding-top:1.2rem;padding-bottom:1rem;max-width:1220px;}
div[data-testid="stVerticalBlock"]{gap:.65rem;}
.stMarkdown p{margin-bottom:.35rem;}
.footer{text-align:center;color:#8a8fa3;font-size:.8rem;margin-top:36px;padding:14px 0;border-top:1px solid rgba(108,92,231,.2);}
.hero{background:linear-gradient(120deg,#4d54d8 0%,#008fa8 100%);
    padding:20px 26px;border-radius:12px;margin-bottom:14px;
      box-shadow:0 8px 24px rgba(108,92,231,.25);}
.hero h1{color:#fff !important;margin:0;font-size:1.8rem;padding:0;}
.hero p{color:rgba(255,255,255,.92) !important;margin:8px 0 0;font-size:1.02rem;}
.steps{display:flex;gap:10px;flex-wrap:wrap;margin-top:14px;}
.steps span{background:rgba(255,255,255,.2);color:#fff;padding:5px 14px;
            border-radius:999px;font-size:.85rem;}
.opt{display:block;padding:5px 10px;margin:3px 0 3px 18px;border-radius:8px;
     background:rgba(108,92,231,.10);border:1px solid rgba(108,92,231,.25);}
.opt b{color:#6C5CE7;margin-right:6px;}
.tag{display:inline-block;padding:2px 12px;border-radius:999px;font-size:.78rem;
     background:rgba(0,184,217,.15);color:#00A3C4;font-weight:600;margin-right:6px;}
div[data-testid="stMetric"]{background:rgba(108,92,231,.08);
     border:1px solid rgba(108,92,231,.2);padding:12px 16px;border-radius:14px;}
.stButton>button,.stDownloadButton>button{width:100%;border-radius:12px;font-weight:600;}
.stDownloadButton>button{background:linear-gradient(135deg,#6C5CE7,#00B8D9);
     color:#fff;border:none;padding:.7rem 1rem;}
section[data-testid="stSidebar"]{border-right:1px solid rgba(108,92,231,.2);}
</style>
"""

# ==========================================================================
# PART 5 : UI (Streamlit)
# ==========================================================================
st.set_page_config(page_title="Classroom Quiz & Notes", page_icon="📝", layout="wide")
st.markdown(CSS, unsafe_allow_html=True)
st.markdown(
    """<div class="hero">
    <h1>📝 Classroom Quiz &amp; Notes</h1>
    <p>One place for teachers to create quizzes and notes, and for students to learn and submit attempts.</p>
    <div class="steps"><span>① Create</span><span>② Share</span>
    <span>③ Attempt</span><span>④ Review results</span></div></div>""",
    unsafe_allow_html=True,
)

if "questions" not in st.session_state:
    st.session_state.questions = []
    st.session_state.meta = {}
if "teacher_user" not in st.session_state:
    st.session_state.teacher_user = None
if "access_role" not in st.session_state:
    st.session_state.access_role = "teacher"

try:
    init_db()
    database_ready = True
except Exception:
    database_ready = False

st.sidebar.markdown("### Access mode")
role_col1, role_col2 = st.sidebar.columns(2)
if role_col1.button("Teacher", use_container_width=True, type="primary" if st.session_state.access_role == "teacher" else "secondary"):
    set_access_role("teacher")
if role_col2.button("Student", use_container_width=True, type="primary" if st.session_state.access_role == "student" else "secondary"):
    set_access_role("student")

if st.session_state.access_role == "teacher":
    with st.container():
        st.caption("Teacher view is active. Student panel is hidden.")

    teacher_tab = st.container()
    with teacher_tab:
        minimum_password_length = 1
        teacher = st.session_state.teacher_user

        if not database_ready:
            st.error("The database is unavailable. Check DATABASE_URL.")
        elif not teacher:
            auth_mode = st.radio("Teacher account", ["Log in", "Create account"], horizontal=True)
            with st.form("teacher_auth_form"):
                username = st.text_input("Username", max_chars=40)
                password = st.text_input("Password", type="password")
                confirm_password = ""
                if auth_mode == "Create account":
                    confirm_password = st.text_input("Confirm password", type="password")
                auth_submitted = st.form_submit_button(auth_mode)

            if auth_submitted:
                if auth_mode == "Create account":
                    if password != confirm_password:
                        st.error("The passwords do not match.")
                    else:
                        try:
                            teacher = create_teacher(username, password, minimum_password_length)
                            st.session_state.teacher_user = teacher
                            st.session_state.questions = []
                            st.session_state.meta = {}
                            st.rerun()
                        except DuplicateUsernameError:
                            st.error("That username is already taken.")
                        except ValueError as error:
                            st.error(str(error))
                else:
                    teacher = authenticate_teacher(username, password)
                    if teacher:
                        st.session_state.teacher_user = teacher
                        st.session_state.questions = []
                        st.session_state.meta = {}
                        st.rerun()
                    st.error("Username or password is incorrect.")

            st.caption("Choose any password length. Password cannot be empty.")

        else:
            identity_col, logout_col = st.columns([5, 1])
            identity_col.caption(f"Signed in as **{teacher['username']}**")
            if logout_col.button("Log out", key="teacher_logout"):
                st.session_state.teacher_user = None
                st.session_state.questions = []
                st.session_state.meta = {}
                st.session_state.pop("published_quiz_code", None)
                st.rerun()

            create_tab, history_tab, notes_tab = st.tabs(["Create Quiz", "Quiz History & OMR", "Notes"])
            with create_tab:
                with st.form("generate_quiz_form"):
                    topic = st.text_input("Topic", placeholder="e.g. Python Loops, DBMS, Data Structures")
                    q_type = st.selectbox("Question type", QUESTION_TYPES)
                    count = st.slider("Number of questions", 5, 50, 10)
                    difficulty = st.selectbox("Difficulty", ["Easy", "Medium", "Hard", "Mixed"], index=1)
                    extra = st.text_area("Extra instructions (optional)",
                                         placeholder="e.g. exam-oriented, numerical-based")
                    generate = st.form_submit_button("Generate Questions")

                if generate:
                    api_key = get_api_key()
                    if not api_key or has_placeholder_key(api_key):
                        st.error("AI generation is not configured. The administrator must set GEMINI_API_KEY on the server.")
                    elif not topic.strip():
                        st.error("Please enter a topic.")
                    else:
                        with st.spinner("Generating questions..."):
                            try:
                                generated = generate_questions(
                                    api_key, DEFAULT_MODEL, topic.strip(), count, q_type, difficulty, extra
                                )
                                if not generated:
                                    st.error("No valid questions were generated. Please try again.")
                                else:
                                    for key in list(st.session_state):
                                        if key.startswith("teacher_sel_"):
                                            del st.session_state[key]
                                    st.session_state.questions = generated
                                    st.session_state.meta = {
                                        "topic": topic.strip(),
                                        "type": q_type,
                                        "difficulty": difficulty,
                                    }
                                    st.session_state.quiz_title = f"{topic.strip()} Quiz"
                                    for index in range(len(generated)):
                                        st.session_state[f"teacher_sel_{index}"] = True
                                    st.session_state.pop("published_quiz_code", None)
                            except json.JSONDecodeError:
                                st.error("The AI returned an unexpected format. Please generate again.")
                            except Exception as error:
                                message = str(error).lower()
                                if "api key" in message or "api_key" in message or "invalid_argument" in message:
                                    st.error("The Gemini API key is invalid or expired. Update the server secret.")
                                elif "unavailable" in message or "503" in message or "rate limit" in message or "429" in message:
                                    st.error("All AI models are currently busy. Wait a little and try again.")
                                else:
                                    st.error("Question generation failed. Check the server logs and try again.")

                questions = st.session_state.questions
                meta = st.session_state.meta
                if questions:
                    st.caption(f"{meta['topic']} · {meta['type']} · {meta['difficulty']}")
                    select_all_col, clear_all_col, first_count_col, first_count_button_col = st.columns(4)
                    select_all_col.button("Select All", on_click=set_all, args=(True,))
                    clear_all_col.button("Clear All", on_click=set_all, args=(False,))
                    first_count = first_count_col.number_input(
                        "First N", min_value=1, max_value=len(questions), value=min(10, len(questions))
                    )
                    first_count_button_col.button("Select First N", on_click=pick_first, args=(first_count,))
                    question_columns = st.columns(2)
                    for index, question in enumerate(questions):
                        with question_columns[index % 2]:
                            with st.container(border=True):
                                st.checkbox(f"Q{index + 1}. {question['question']}", key=f"teacher_sel_{index}")
                                if meta["type"] == "MCQ":
                                    for letter, option in zip(LETTERS, question["options"]):
                                        st.write(f"({letter}) {option}")
                                    with st.expander("Answer key"):
                                        answer_index = LETTERS.index(question["answer"])
                                        st.write(f"Correct: ({question['answer']}) {question['options'][answer_index]}")
                                        if question.get("explanation"):
                                            st.caption(question["explanation"])
                                else:
                                    with st.expander("Model answer"):
                                        st.write(question.get("answer", ""))

                    chosen = selected_questions()
                    st.caption(f"{len(chosen)} selected")
                    pdf_title = st.text_input("Quiz title", key="quiz_title")
                    include_answers = st.checkbox("Include answer key in teacher PDF", value=True)
                    if chosen:
                        pdf_bytes = build_pdf(pdf_title, chosen, meta["type"], include_answers)
                        st.download_button(
                            "Download PDF",
                            data=pdf_bytes,
                            file_name=f"{'_'.join(meta['topic'].split())}_quiz.pdf",
                            mime="application/pdf",
                        )

                    if meta["type"] != "MCQ":
                        st.info("Automatic attempts and marks are currently supported for MCQ quizzes.")
                    elif not chosen:
                        st.warning("Select at least one question before publishing.")
                    elif st.button("Publish Quiz", type="primary"):
                        try:
                            code = create_quiz(
                                teacher["id"], pdf_title or f"{meta['topic']} Quiz", meta["topic"], meta["difficulty"], chosen
                            )
                            st.session_state.published_quiz_code = code
                        except Exception:
                            st.error("Could not save this quiz. Check the database connection.")

                    if st.session_state.get("published_quiz_code"):
                        st.success("Quiz is live. Students can find it in their panel:")
                        st.code(st.session_state.published_quiz_code)

            with history_tab:
                try:
                    quizzes = list_quizzes(teacher["id"])
                except Exception:
                    quizzes = []
                    st.error("Could not load your quiz history.")

                if quizzes:
                    st.dataframe([
                        {
                            "Quiz": quiz["title"],
                            "Code": quiz["code"],
                            "Topic": quiz["topic"],
                            "Created": quiz["created_at"].isoformat(sep=" ", timespec="minutes"),
                        }
                        for quiz in quizzes
                    ], width="stretch", hide_index=True)
                    selected_code = st.selectbox(
                        "Quiz results",
                        [quiz["code"] for quiz in quizzes],
                        format_func=lambda code: next(
                            f"{quiz['title']} · {code}" for quiz in quizzes if quiz["code"] == code
                        ),
                    )
                    attempts = get_attempts(selected_code, teacher["id"])
                    if attempts:
                        quiz_title = next(quiz["title"] for quiz in quizzes if quiz["code"] == selected_code)
                        st.download_button(
                            "Download Excel (.xlsx)",
                            data=build_marksheet_xlsx(quiz_title, attempts),
                            file_name=f"{selected_code}_marks.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                        st.dataframe([
                            {
                                "UID": attempt["student_uid"],
                                "Name": attempt["student_name"],
                                "Marks": f"{attempt['score']} / {attempt['total']}",
                                "Percentage": f"{attempt['percentage']}%",
                                "Submitted": attempt["submitted_at"].isoformat(sep=" ", timespec="minutes"),
                            }
                            for attempt in attempts
                        ], width="stretch", hide_index=True)
                        selected_attempt_id = st.selectbox(
                            "Review OMR sheet",
                            [attempt["id"] for attempt in attempts],
                            format_func=lambda attempt_id: next(
                                f"{attempt['student_name']} · {attempt['student_uid']} · "
                                f"{attempt['submitted_at'].isoformat(sep=' ', timespec='minutes')}"
                                for attempt in attempts if attempt["id"] == attempt_id
                            ),
                        )
                        review = get_attempt_review(selected_attempt_id, teacher["id"])
                        if review:
                            st.metric("Student score", f"{review['score']} / {review['total']}")
                            for index, question in enumerate(review["questions"]):
                                chosen_answer = review["answers"].get(str(index), "")
                                correct_answer = question["answer"]
                                st.write(f"Q{index + 1}. {question['question']}")
                                st.caption(f"Student marked: {chosen_answer or 'Blank'} · Correct: {correct_answer}")
                                if chosen_answer == correct_answer:
                                    st.success("Correct")
                                else:
                                    st.error("Incorrect")
                    else:
                        st.info("No student attempts for this quiz yet.")
                else:
                    st.info("Your published quizzes will appear here with their dates and results.")

            with notes_tab:
                with st.form("create_note_form"):
                    note_title = st.text_input("Note title", max_chars=200)
                    note_content = st.text_area("Notes", height=180, max_chars=50_000)
                    note_saved = st.form_submit_button("Save and Share Notes")
                if note_saved:
                    try:
                        st.session_state.shared_note_code = create_note(teacher["id"], note_title, note_content)
                    except ValueError as error:
                        st.error(str(error))
                    except Exception:
                        st.error("Could not save these notes. Check the database connection.")
                if st.session_state.get("shared_note_code"):
                    st.success("Notes saved. They are now available in the student panel:")
                    st.code(st.session_state.shared_note_code)

                teacher_notes = list_notes(teacher["id"])
                if teacher_notes:
                    st.dataframe([
                        {
                            "Title": note["title"],
                            "Share code": note["code"],
                            "Created": note["created_at"].isoformat(sep=" ", timespec="minutes"),
                        }
                        for note in teacher_notes
                    ], width="stretch", hide_index=True)

            cheating_events = list_cheating_events(teacher["id"])
            if cheating_events:
                st.warning("Security alert: unauthorized access detected on the student side.")
                st.dataframe([
                    {
                        "Student": f"{event['student_name']} ({event['student_uid']})",
                        "Attempted path": event["attempted_path"],
                        "Details": event["details"] or "No detail provided",
                        "Time": event["recorded_at"].isoformat(sep=" ", timespec="seconds"),
                    }
                    for event in cheating_events
                ], width="stretch", hide_index=True)

else:
    with st.container():
        st.caption("Student view is active. Teacher dashboard is locked.")
        st.warning("Only the student panel is allowed in this mode. If someone tries to open the teacher area, the action will be logged as cheating.")

        if st.sidebar.button("Teacher panel (blocked)", key="blocked_teacher_panel"):
            student_uid = st.session_state.get("student_uid_active", "UNKNOWN")
            student_name = st.session_state.get("student_name_active", "Unknown student")
            record_cheating_event(
                student_uid=student_uid,
                student_name=student_name,
                attempted_path="teacher_panel",
                details="Student clicked the teacher-only panel while in student mode.",
                teacher_id=st.session_state.teacher_user["id"] if st.session_state.teacher_user else None,
            )
            st.error("Unauthorized teacher-panel access detected. This attempt has been reported to the teacher.")
            st.stop()

    student_tab = st.container()
    with student_tab:
        if not database_ready:
            st.error("The database is unavailable. Ask the app administrator to check the database.")
        else:
            quiz_panel, notes_panel = st.tabs(["Quizzes", "Shared Notes"])
            with quiz_panel:
                try:
                    available_quizzes = list_published_quizzes()
                except Exception:
                    available_quizzes = []
                    st.error("Could not load published quizzes.")

                if available_quizzes:
                    selected_quiz_code = st.selectbox(
                        "Published quizzes",
                        [quiz["code"] for quiz in available_quizzes],
                        format_func=lambda code: next(
                            f"{quiz['title']} · {quiz['created_at'].strftime('%Y-%m-%d')}"
                            for quiz in available_quizzes if quiz["code"] == code
                        ),
                    )
                    st.caption("Quiz code")
                    st.code(selected_quiz_code)
                    if st.button("Open selected quiz", type="primary"):
                        st.session_state.active_quiz_code = selected_quiz_code
                else:
                    selected_quiz_code = ""
                    st.info("Published quizzes will appear here.")

                active_code = st.session_state.get("active_quiz_code", "")
                if active_code:
                    quiz = get_quiz(active_code)
                    if quiz:
                        st.markdown(f"### {html_escape(quiz['title'])}")
                        st.caption(f"{quiz['topic']} · {quiz['difficulty']} · {len(quiz['questions'])} questions")
                        uid_col, name_col = st.columns(2)
                        student_uid = uid_col.text_input("Student UID", key=f"student_uid_{active_code}", max_chars=100)
                        student_name = name_col.text_input("Student name", key=f"student_name_{active_code}", max_chars=200)
                        st.session_state.student_uid_active = student_uid
                        st.session_state.student_name_active = student_name
                        previous_result = get_student_result(active_code, student_uid)

                        if previous_result:
                            st.success(
                                f"Submitted {previous_result['submitted_at'].isoformat(sep=' ', timespec='minutes')}"
                            )
                            st.metric("Your marks", f"{previous_result['score']} / {previous_result['total']}")
                        else:
                            answers = {}
                            questions_column, progress_column = st.columns([3, 1])
                            with questions_column:
                                question_columns = st.columns(2)
                                for index, question in enumerate(quiz["questions"]):
                                    with question_columns[index % 2]:
                                        answers[str(index)] = st.radio(
                                            f"Q{index + 1}. {question['question']}",
                                            options=LETTERS,
                                            format_func=lambda letter, options=question["options"]: (
                                                f"({letter}) {options[LETTERS.index(letter)]}"
                                            ),
                                            index=None,
                                            key=f"student_answer_{active_code}_{index}",
                                        ) or ""

                            answered_count = sum(bool(answer) for answer in answers.values())
                            remaining_count = len(quiz["questions"]) - answered_count
                            with progress_column:
                                with st.container(border=True):
                                    st.metric("Answered", f"{answered_count} / {len(quiz['questions'])}")
                                    st.metric("Remaining", remaining_count)
                                    st.progress(answered_count / len(quiz["questions"]))

                            if st.button("Submit Quiz", type="primary", key=f"submit_quiz_{active_code}"):
                                if not student_uid.strip() or not student_name.strip():
                                    st.error("Enter both your student UID and name.")
                                elif remaining_count:
                                    st.error("Answer every question before submitting.")
                                else:
                                    try:
                                        submit_attempt(active_code, student_uid, student_name, answers)
                                        st.rerun()
                                    except DuplicateAttemptError:
                                        st.error("This UID has already submitted this quiz.")
                                    except Exception:
                                        st.error("Your attempt could not be saved. Please try again.")

            with notes_panel:
                try:
                    available_notes = list_shared_notes()
                except Exception:
                    available_notes = []
                    st.error("Could not load shared notes.")
                if available_notes:
                    selected_note_code = st.selectbox(
                        "Shared notes",
                        [note["code"] for note in available_notes],
                        format_func=lambda code: next(
                            f"{note['title']} · {note['created_at'].strftime('%Y-%m-%d')}"
                            for note in available_notes if note["code"] == code
                        ),
                    )
                    st.caption("Notes code")
                    st.code(selected_note_code)
                    selected_note = get_note(selected_note_code)
                    if selected_note:
                        st.markdown(f"### {html_escape(selected_note['title'])}")
                        st.markdown(selected_note["content"])
                else:
                    st.info("Teacher-shared notes will appear here.")

st.markdown('<div class="footer">Classroom Quiz & Notes · Built with Python, Streamlit and Gemini</div>', unsafe_allow_html=True)