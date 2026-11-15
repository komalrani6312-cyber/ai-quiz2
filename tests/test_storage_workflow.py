import io
import os
import sqlite3
import tempfile
import unittest
import uuid
import zipfile
from pathlib import Path

os.environ["DATABASE_URL"] = "sqlite://"

import storage
from openpyxl import load_workbook
from sqlalchemy import URL, create_engine, inspect
from sqlalchemy.orm import sessionmaker


class StorageWorkflowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        storage.init_db()

    def setUp(self):
        unique = uuid.uuid4().hex[:8]
        self.teacher = storage.create_teacher(f"teacher_{unique}", "localpass", 4)
        self.other_teacher = storage.create_teacher(f"other_{unique}", "otherpass", 4)
        self.questions = [{
            "question": "Which number is even?",
            "options": ["1", "2", "3", "5"],
            "answer": "B",
            "explanation": "2 is even.",
        }]

    def test_quiz_attempt_review_and_teacher_isolation(self):
        code = storage.create_quiz(
            self.teacher["id"], "Numbers", "Arithmetic", "Easy", self.questions
        )
        result = storage.submit_attempt(code, "student-1", "Student One", {"0": "B"})

        self.assertEqual(result, {"score": 1, "total": 1})
        attempts = storage.get_attempts(code, self.teacher["id"])
        self.assertEqual(len(attempts), 1)
        self.assertIsNone(storage.get_attempt_review(attempts[0]["id"], self.other_teacher["id"]))

        review = storage.get_attempt_review(attempts[0]["id"], self.teacher["id"])
        self.assertEqual(review["answers"], {"0": "B"})
        self.assertEqual(review["questions"][0]["answer"], "B")
        self.assertEqual(storage.get_student_result(code, "student-1")["score"], 1)
        self.assertEqual(storage.list_quizzes(self.other_teacher["id"]), [])

        with self.assertRaises(storage.DuplicateAttemptError):
            storage.submit_attempt(code, "STUDENT-1", "Student One", {"0": "A"})

        workbook_bytes = storage.build_marksheet_xlsx("Numbers", attempts)
        self.assertTrue(zipfile.is_zipfile(io.BytesIO(workbook_bytes)))
        worksheet = load_workbook(io.BytesIO(workbook_bytes), data_only=True).active
        self.assertEqual(worksheet["B2"].value, "STUDENT-1")
        self.assertEqual(worksheet["D2"].value, 1)

    def test_notes_are_shared_and_owned(self):
        code = storage.create_note(self.teacher["id"], "Revision", "Chapter notes")
        self.assertEqual(storage.get_note(code)["content"], "Chapter notes")
        self.assertEqual(storage.list_notes(self.other_teacher["id"]), [])
        self.assertEqual(storage.list_shared_notes()[0]["code"], code)

    def test_cheating_alerts_are_recorded_and_visible_to_teacher(self):
        event = storage.record_cheating_event(
            student_uid="STU-999",
            student_name="Test Student",
            quiz_code="QZ12345",
            attempted_path="teacher_panel",
            teacher_id=self.teacher["id"],
            details="Student clicked teacher-only area",
        )
        self.assertEqual(event["student_uid"], "STU-999")
        self.assertEqual(event["attempted_path"], "teacher_panel")
        alerts = storage.list_cheating_events(self.teacher["id"])
        self.assertTrue(any(alert["student_uid"] == "STU-999" for alert in alerts))
        self.assertEqual(storage.list_cheating_events(self.other_teacher["id"]), [])

    def test_teacher_password_is_hashed_and_verified(self):
        self.assertIsNotNone(storage.authenticate_teacher(self.teacher["username"], "localpass"))
        self.assertIsNone(storage.authenticate_teacher(self.teacher["username"], "wrongpass"))

    def test_existing_database_migrates_and_keeps_quiz(self):
        temporary_file = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
        database_path = Path(temporary_file.name)
        temporary_file.close()
        connection = sqlite3.connect(database_path)
        connection.executescript(
            "CREATE TABLE quizzes (code VARCHAR(12) PRIMARY KEY, title VARCHAR(200) NOT NULL, "
            "topic VARCHAR(200) NOT NULL, difficulty VARCHAR(30) NOT NULL, questions_json TEXT NOT NULL, "
            "created_at DATETIME NOT NULL);"
            "CREATE TABLE attempts (id INTEGER PRIMARY KEY AUTOINCREMENT, quiz_code VARCHAR(12) NOT NULL, "
            "student_uid VARCHAR(100) NOT NULL, student_name VARCHAR(200) NOT NULL, score INTEGER NOT NULL, "
            "total INTEGER NOT NULL, submitted_at DATETIME NOT NULL, "
            "CONSTRAINT uq_attempt_quiz_student UNIQUE (quiz_code, student_uid));"
        )
        connection.execute(
            "INSERT INTO quizzes VALUES (?, ?, ?, ?, ?, ?)",
            (
                "OLD12345",
                "Existing quiz",
                "Legacy",
                "Easy",
                '[{"question":"Old question?","options":["a","b","c","d"],"answer":"A"}]',
                "2026-01-01 00:00:00",
            ),
        )
        connection.commit()
        connection.close()

        old_engine = storage.engine
        old_factory = storage.SessionFactory
        migrated_engine = create_engine(URL.create("sqlite", database=str(database_path)))
        storage.engine = migrated_engine
        storage.SessionFactory = sessionmaker(bind=migrated_engine, expire_on_commit=False)
        try:
            storage.init_db()
            columns = {column["name"] for column in inspect(migrated_engine).get_columns("attempts")}
            self.assertIn("answers_json", columns)
            teacher = storage.create_teacher(f"legacy_{uuid.uuid4().hex[:8]}", "localpass", 4)
            self.assertEqual(storage.list_quizzes(teacher["id"])[0]["code"], "OLD12345")
        finally:
            storage.engine = old_engine
            storage.SessionFactory = old_factory
            migrated_engine.dispose()
            database_path.unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
