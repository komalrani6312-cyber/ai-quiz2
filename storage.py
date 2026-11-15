import json
import hashlib
import os
import re
import secrets
import hmac
from datetime import datetime, timezone
from io import BytesIO

from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font
from sqlalchemy import DateTime, ForeignKey, Index, Integer, String, Text, UniqueConstraint, create_engine, inspect, select, text
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, sessionmaker

load_dotenv()

database_url = (os.getenv("DATABASE_URL") or "sqlite:///quiz_app.db").strip()
if database_url.startswith("postgres://"):
    database_url = database_url.replace("postgres://", "postgresql+psycopg://", 1)
elif database_url.startswith("postgresql://"):
    database_url = database_url.replace("postgresql://", "postgresql+psycopg://", 1)

engine_options = {"pool_pre_ping": True}
if database_url.startswith("sqlite:"):
    engine_options["connect_args"] = {"check_same_thread": False}

engine = create_engine(database_url, **engine_options)
SessionFactory = sessionmaker(bind=engine, expire_on_commit=False)


class Base(DeclarativeBase):
    pass


class Teacher(Base):
    __tablename__ = "teachers"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    username: Mapped[str] = mapped_column(String(40), unique=True, nullable=False, index=True)
    password_hash: Mapped[str] = mapped_column(String(200), nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


class Quiz(Base):
    __tablename__ = "quizzes"

    code: Mapped[str] = mapped_column(String(12), primary_key=True)
    title: Mapped[str] = mapped_column(String(200), nullable=False)
    topic: Mapped[str] = mapped_column(String(200), nullable=False)
    difficulty: Mapped[str] = mapped_column(String(30), nullable=False)
    questions_json: Mapped[str] = mapped_column(Text, nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    teacher_id: Mapped[int | None] = mapped_column(ForeignKey("teachers.id"), nullable=True, index=True)


class Attempt(Base):
    __tablename__ = "attempts"
    __table_args__ = (UniqueConstraint("quiz_code", "student_uid", name="uq_attempt_quiz_student"),)

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    quiz_code: Mapped[str] = mapped_column(ForeignKey("quizzes.code"), nullable=False, index=True)
    student_uid: Mapped[str] = mapped_column(String(100), nullable=False)
    student_name: Mapped[str] = mapped_column(String(200), nullable=False)
    score: Mapped[int] = mapped_column(Integer, nullable=False)
    total: Mapped[int] = mapped_column(Integer, nullable=False)
    answers_json: Mapped[str] = mapped_column(Text, nullable=False, default="{}")
    submitted_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


class CheatingEvent(Base):
    __tablename__ = "cheating_events"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    teacher_id: Mapped[int | None] = mapped_column(ForeignKey("teachers.id"), nullable=True, index=True)
    student_uid: Mapped[str] = mapped_column(String(100), nullable=False, default="UNKNOWN")
    student_name: Mapped[str] = mapped_column(String(200), nullable=False, default="Unknown student")
    quiz_code: Mapped[str | None] = mapped_column(String(12), nullable=True, index=True)
    attempted_path: Mapped[str] = mapped_column(String(100), nullable=False, default="unknown")
    details: Mapped[str] = mapped_column(Text, nullable=False, default="")
    recorded_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


class Note(Base):
    __tablename__ = "notes"

    code: Mapped[str] = mapped_column(String(12), primary_key=True)
    teacher_id: Mapped[int] = mapped_column(ForeignKey("teachers.id"), nullable=False, index=True)
    title: Mapped[str] = mapped_column(String(200), nullable=False)
    content: Mapped[str] = mapped_column(Text, nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


class DuplicateAttemptError(Exception):
    pass


class DuplicateUsernameError(Exception):
    pass


def init_db():
    Base.metadata.create_all(engine)
    with engine.begin() as connection:
        quiz_columns = {column["name"] for column in inspect(connection).get_columns("quizzes")}
        if "teacher_id" not in quiz_columns:
            connection.execute(text("ALTER TABLE quizzes ADD COLUMN teacher_id INTEGER REFERENCES teachers(id)"))

        attempt_columns = {column["name"] for column in inspect(connection).get_columns("attempts")}
        if "answers_json" not in attempt_columns:
            connection.execute(text("ALTER TABLE attempts ADD COLUMN answers_json TEXT NOT NULL DEFAULT '{}'"))

    Index("ix_quizzes_teacher_id", Quiz.teacher_id).create(engine, checkfirst=True)


def record_cheating_event(student_uid, student_name, quiz_code=None, attempted_path="unknown", teacher_id=None, details=""):
    normalized_uid = (student_uid or "").strip().upper() or "UNKNOWN"
    cleaned_name = " ".join((student_name or "").split()) or "Unknown student"
    event = CheatingEvent(
        teacher_id=teacher_id,
        student_uid=normalized_uid,
        student_name=cleaned_name,
        quiz_code=(quiz_code or "").strip().upper() or None,
        attempted_path=(attempted_path or "unknown").strip()[:100] or "unknown",
        details=(details or "").strip()[:5000],
        recorded_at=datetime.now(timezone.utc),
    )
    with SessionFactory.begin() as session:
        session.add(event)
        session.flush()
    return {
        "id": event.id,
        "teacher_id": event.teacher_id,
        "student_uid": event.student_uid,
        "student_name": event.student_name,
        "quiz_code": event.quiz_code,
        "attempted_path": event.attempted_path,
        "details": event.details,
        "recorded_at": event.recorded_at,
    }


def list_cheating_events(teacher_id):
    with SessionFactory() as session:
        events = session.scalars(
            select(CheatingEvent)
            .where(CheatingEvent.teacher_id == teacher_id)
            .order_by(CheatingEvent.recorded_at.desc())
        ).all()
        return [
            {
                "id": event.id,
                "student_uid": event.student_uid,
                "student_name": event.student_name,
                "quiz_code": event.quiz_code,
                "attempted_path": event.attempted_path,
                "details": event.details,
                "recorded_at": event.recorded_at,
            }
            for event in events
        ]


def _password_hash(password, salt=None):
    salt_bytes = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt_bytes, 310_000)
    return f"pbkdf2_sha256${salt_bytes.hex()}${digest.hex()}"


def create_teacher(username, password, minimum_password_length=8):
    clean_username = username.strip().lower()
    if not re.fullmatch(r"[a-z0-9_.-]{3,40}", clean_username):
        raise ValueError("Username must be 3-40 characters using letters, numbers, dot, dash, or underscore.")
    if len(password) < minimum_password_length:
        raise ValueError(f"Password must be at least {minimum_password_length} characters.")

    teacher = Teacher(
        username=clean_username,
        password_hash=_password_hash(password),
        created_at=datetime.now(timezone.utc),
    )
    try:
        with SessionFactory.begin() as session:
            session.add(teacher)
            session.flush()
            session.query(Quiz).filter(Quiz.teacher_id.is_(None)).update(
                {Quiz.teacher_id: teacher.id}, synchronize_session=False
            )
        return {"id": teacher.id, "username": teacher.username}
    except IntegrityError as error:
        raise DuplicateUsernameError("That username is already registered.") from error


def authenticate_teacher(username, password):
    with SessionFactory() as session:
        teacher = session.scalar(select(Teacher).where(Teacher.username == username.strip().lower()))
        if teacher is None:
            return None
        try:
            algorithm, salt_hex, expected_hex = teacher.password_hash.split("$", 2)
            if algorithm != "pbkdf2_sha256":
                return None
            actual_hash = _password_hash(password, bytes.fromhex(salt_hex)).split("$", 2)[2]
        except (ValueError, TypeError):
            return None
        if not hmac.compare_digest(actual_hash, expected_hex):
            return None
        return {"id": teacher.id, "username": teacher.username}


def create_quiz(teacher_id, title, topic, difficulty, questions):
    cleaned_questions = [
        {
            "question": str(question["question"]),
            "options": [str(option) for option in question["options"]],
            "answer": str(question["answer"]).upper(),
            "explanation": str(question.get("explanation", "")),
        }
        for question in questions
    ]
    if not cleaned_questions or any(
        len(question["options"]) != 4 or question["answer"] not in {"A", "B", "C", "D"}
        for question in cleaned_questions
    ):
        raise ValueError("A quiz must contain valid multiple-choice questions.")

    for _ in range(5):
        code = "".join(secrets.choice("23456789ABCDEFGHJKLMNPQRSTUVWXYZ") for _ in range(8))
        quiz = Quiz(
            code=code,
            title=title.strip()[:200],
            topic=topic.strip()[:200],
            difficulty=difficulty[:30],
            questions_json=json.dumps(cleaned_questions),
            created_at=datetime.now(timezone.utc),
            teacher_id=teacher_id,
        )
        try:
            with SessionFactory.begin() as session:
                session.add(quiz)
            return code
        except IntegrityError:
            continue
    raise RuntimeError("Could not create a unique quiz code. Please try again.")


def get_quiz(code):
    with SessionFactory() as session:
        quiz = session.get(Quiz, code.strip().upper())
        if quiz is None:
            return None
        return {
            "code": quiz.code,
            "title": quiz.title,
            "topic": quiz.topic,
            "difficulty": quiz.difficulty,
            "questions": json.loads(quiz.questions_json),
            "created_at": quiz.created_at,
        }


def list_quizzes(teacher_id):
    with SessionFactory() as session:
        quizzes = session.scalars(
            select(Quiz).where(Quiz.teacher_id == teacher_id).order_by(Quiz.created_at.desc())
        ).all()
        return [
            {"code": quiz.code, "title": quiz.title, "topic": quiz.topic, "created_at": quiz.created_at}
            for quiz in quizzes
        ]


def list_published_quizzes():
    with SessionFactory() as session:
        quizzes = session.scalars(select(Quiz).order_by(Quiz.created_at.desc())).all()
        return [
            {"code": quiz.code, "title": quiz.title, "topic": quiz.topic, "created_at": quiz.created_at}
            for quiz in quizzes
        ]


def submit_attempt(code, student_uid, student_name, answers):
    normalized_code = code.strip().upper()
    normalized_uid = student_uid.strip().upper()
    cleaned_name = " ".join(student_name.split())
    if not normalized_uid or not cleaned_name:
        raise ValueError("Enter both your student UID and name.")
    if len(normalized_uid) > 100 or len(cleaned_name) > 200:
        raise ValueError("UID or name is too long.")

    try:
        with SessionFactory.begin() as session:
            quiz = session.get(Quiz, normalized_code)
            if quiz is None:
                raise LookupError("Quiz code was not found.")
            questions = json.loads(quiz.questions_json)
            normalized_answers = {
                str(index): answer if answer in {"A", "B", "C", "D"} else ""
                for index, answer in answers.items()
                if str(index).isdigit() and 0 <= int(index) < len(questions)
            }
            score = sum(
                normalized_answers.get(str(index)) == question["answer"]
                for index, question in enumerate(questions)
            )
            attempt = Attempt(
                quiz_code=quiz.code,
                student_uid=normalized_uid,
                student_name=cleaned_name,
                score=score,
                total=len(questions),
                answers_json=json.dumps(normalized_answers),
                submitted_at=datetime.now(timezone.utc),
            )
            session.add(attempt)
        return {"score": score, "total": len(questions)}
    except IntegrityError as error:
        raise DuplicateAttemptError("This UID has already submitted this quiz.") from error


def get_attempts(code, teacher_id):
    with SessionFactory() as session:
        quiz = session.get(Quiz, code.strip().upper())
        if quiz is None or quiz.teacher_id != teacher_id:
            return []
        attempts = session.scalars(
            select(Attempt)
            .where(Attempt.quiz_code == code.strip().upper())
            .order_by(Attempt.submitted_at.desc(), Attempt.student_uid.asc())
        ).all()
        return [
            {
                "id": attempt.id,
                "student_uid": attempt.student_uid,
                "student_name": attempt.student_name,
                "score": attempt.score,
                "total": attempt.total,
                "percentage": round(attempt.score * 100 / attempt.total, 1) if attempt.total else 0,
                "submitted_at": attempt.submitted_at,
            }
            for attempt in attempts
        ]


def get_attempt_review(attempt_id, teacher_id):
    with SessionFactory() as session:
        attempt = session.get(Attempt, attempt_id)
        if attempt is None:
            return None
        quiz = session.get(Quiz, attempt.quiz_code)
        if quiz is None or quiz.teacher_id != teacher_id:
            return None
        return {
            "quiz_title": quiz.title,
            "student_uid": attempt.student_uid,
            "student_name": attempt.student_name,
            "score": attempt.score,
            "total": attempt.total,
            "answers": json.loads(attempt.answers_json or "{}"),
            "questions": json.loads(quiz.questions_json),
        }


def create_note(teacher_id, title, content):
    cleaned_title = " ".join(title.split())
    cleaned_content = content.strip()
    if not cleaned_title or not cleaned_content:
        raise ValueError("Enter both a note title and its content.")
    if len(cleaned_title) > 200 or len(cleaned_content) > 50_000:
        raise ValueError("The note title or content is too long.")

    for _ in range(5):
        code = "".join(secrets.choice("23456789ABCDEFGHJKLMNPQRSTUVWXYZ") for _ in range(8))
        note = Note(
            code=code,
            teacher_id=teacher_id,
            title=cleaned_title,
            content=cleaned_content,
            created_at=datetime.now(timezone.utc),
        )
        try:
            with SessionFactory.begin() as session:
                session.add(note)
            return code
        except IntegrityError:
            continue
    raise RuntimeError("Could not create a unique notes code. Please try again.")


def list_notes(teacher_id):
    with SessionFactory() as session:
        notes = session.scalars(
            select(Note).where(Note.teacher_id == teacher_id).order_by(Note.created_at.desc())
        ).all()
        return [
            {"code": note.code, "title": note.title, "created_at": note.created_at}
            for note in notes
        ]


def get_note(code):
    with SessionFactory() as session:
        note = session.get(Note, code.strip().upper())
        if note is None:
            return None
        return {"code": note.code, "title": note.title, "content": note.content}


def list_shared_notes():
    with SessionFactory() as session:
        notes = session.scalars(select(Note).order_by(Note.created_at.desc())).all()
        return [
            {"code": note.code, "title": note.title, "created_at": note.created_at}
            for note in notes
        ]


def get_student_result(code, student_uid):
    normalized_code = code.strip().upper()
    normalized_uid = student_uid.strip().upper()
    if not normalized_uid:
        return None
    with SessionFactory() as session:
        attempt = session.scalar(
            select(Attempt).where(
                Attempt.quiz_code == normalized_code,
                Attempt.student_uid == normalized_uid,
            )
        )
        if attempt is None:
            return None
        return {
            "score": attempt.score,
            "total": attempt.total,
            "submitted_at": attempt.submitted_at,
        }


def build_marksheet_xlsx(quiz_title, attempts):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Marks"
    sheet.append(["Quiz", "Student UID", "Student Name", "Marks", "Total", "Percentage", "Submitted At"])
    for cell in sheet[1]:
        cell.font = Font(bold=True)

    def safe_text(value):
        text = str(value)
        return f"'{text}" if text.startswith(("=", "+", "-", "@")) else text

    for attempt in attempts:
        submitted_at = attempt["submitted_at"]
        sheet.append([
            safe_text(quiz_title),
            safe_text(attempt["student_uid"]),
            safe_text(attempt["student_name"]),
            attempt["score"],
            attempt["total"],
            attempt["percentage"],
            submitted_at.isoformat(sep=" ", timespec="seconds") if submitted_at else "",
        ])

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column, width in {"A": 30, "B": 20, "C": 28, "D": 12, "E": 12, "F": 16, "G": 24}.items():
        sheet.column_dimensions[column].width = width

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()