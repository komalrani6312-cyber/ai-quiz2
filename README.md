# AI Quiz & Question Generator

Classroom Quiz & Notes is a Streamlit app for teacher accounts, generated quizzes, student attempts, OMR-style answer review, class notes, and Excel marksheets.

## Run locally

1. Install Python 3.10 or later.
2. Install dependencies:

   ```powershell
   pip install -r requirements.txt
   ```

3. Copy `env.example` to `.env` and set `GEMINI_API_KEY`.
   `DATABASE_URL` may be left blank locally; the app then uses a local SQLite database.
4. Start the app:

   ```powershell
   streamlit run app.py
   ```

5. In **Teacher Dashboard**, create a teacher account with a unique username and password. Passwords are stored as salted hashes. The local minimum is 4 characters; public deployments require at least 12 characters.

Never commit `.env` or share its contents. The `env.example` file must contain only a placeholder.

## Deploy on Render

1. Push this project to a GitHub repository. Confirm `.env` is not committed.
2. In Render, choose **New +** then **Blueprint**, and connect the repository. Render reads `render.yaml` and builds the Streamlit service.
3. Create a PostgreSQL database with a hosted provider such as Supabase. Copy its connection URL.
4. In the Render service's Environment settings, set `GEMINI_API_KEY` and `DATABASE_URL`. Keep both out of GitHub. Teachers create their own accounts in the app.
5. Open the Render service URL after deployment completes. The database tables are created automatically.

The API key stays on the server. Students only need a published quiz, their UID, and name. All visitors share the configured Gemini quota, so add usage controls before unrestricted public promotion.

## Data and persistence

Published quizzes, student attempts, OMR responses, and notes are stored in the configured database. A UID can submit once per quiz; marks are calculated server-side. Each teacher sees only their own quiz history, results, OMR reviews, and notes. Students see published quiz/note codes automatically, but never see quiz answer keys. Teachers can download an `.xlsx` marksheet with submission dates. Automatic marks are currently for MCQs; short- and long-answer questions remain available for teacher PDF generation only.

Use a hosted PostgreSQL database for deployment. Local SQLite is intended for development and is not durable storage on most app hosts. The database stores teacher usernames and password hashes, quiz questions and answer keys, student UIDs/names, selected answers, marks, notes, and timestamps. Student UID/name are entered by the student and are not independently verified. Limit database access and retain student records only as long as needed.

## Features

- Generate 5-50 questions with type and difficulty settings
- Teacher accounts with separate quiz and notes history
- Generate quiz questions and publish them to the student panel
- Students attempt a quiz with their UID and name
- Prevent duplicate submissions for the same UID and quiz
- Teacher dashboard with submission history, OMR review, and Excel export
- Create and share class notes
- Teacher PDF export with an optional answer key
