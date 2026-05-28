# TRP Cover Letter Generator

An AI-powered tool for generating **Temporary Resident Permit (TRP)** cover letters for criminally inadmissible foreign nationals applying for temporary entry to Canada. Built for use at Cohen Immigration Law.

## What It Does

Case managers enter client information and criminal history, upload court documents, and the tool produces a professionally formatted, legally sound DOCX cover letter ready to submit to a Canadian consulate or port of entry.

**Key capabilities:**

- Extracts criminal charges from uploaded PDF/Word court documents using Claude AI
- Maps foreign charges (US state statutes, etc.) to Canadian Criminal Code equivalents with date-aware statute history (implements the *Tran* principle)
- Generates TRP cover letters via Claude, tailored to 14+ travel purposes (business, cruise, wedding, flight crew, hunting/fishing, etc.)
- Downloads letters as formatted DOCX with firm letterhead, conviction table, legal defense notes, and attorney signature
- Saves and retrieves letters from a Supabase database, organized by case manager

## Tech Stack

- **Frontend**: Vanilla HTML/CSS/JavaScript (no build step)
- **Backend**: Node.js + Express
- **AI**: Claude Sonnet (`claude-sonnet-4-6`) — document extraction, offence parsing, letter generation
- **Database**: Supabase (PostgreSQL) — persistent letter storage
- **Document I/O**: `pdf-parse`, `mammoth` (DOCX input), `docx` (DOCX output)
- **Hosting**: Vercel

## Prerequisites

- Node.js 18+
- An [Anthropic API key](https://console.anthropic.com/)
- A [Supabase](https://supabase.com/) project

## Local Setup

1. Clone the repo and install dependencies:

   ```bash
   git clone https://github.com/dimitramarks/cover-letter-generator.git
   cd cover-letter-generator
   npm install
   ```

2. Create a `.env` file in the project root:

   ```
   ANTHROPIC_API_KEY=your_anthropic_api_key_here
   TRP_PORT=8081

   SUPABASE_URL=https://your-project-ref.supabase.co
   SUPABASE_ANON_KEY=your_anon_key
   SUPABASE_SERVICE_ROLE_KEY=your_service_role_key
   ```

3. Create the `letters` table in your Supabase project (**SQL Editor → New query**):

   ```sql
   CREATE TABLE letters (
     id              TEXT PRIMARY KEY,
     saved_at        TIMESTAMPTZ NOT NULL,
     case_manager    TEXT,
     client_name     TEXT,
     client_dob      TEXT,
     offence_summary TEXT,
     travel_purpose  TEXT,
     entry_type      TEXT,
     date_from       TEXT,
     date_to         TEXT,
     payload         JSONB NOT NULL,
     generated       JSONB NOT NULL
   );
   ```

4. Start the server:

   ```bash
   npm start
   ```

5. Open [http://localhost:8081](http://localhost:8081) in your browser.

## Vercel Deployment

The project is deployed to Vercel and auto-deploys on every push to `main`.

**Live URL:** https://cover-letter-generator-fawn-one.vercel.app

To configure environment variables in Vercel:
**Dashboard → Project → Settings → Environment Variables** — add the same four keys from the `.env` file above (`TRP_PORT` is not needed on Vercel).

## Usage

1. Fill in client details (name, date of birth, nationality, pronouns)
2. Add criminal history — enter charges manually or upload court documents (PDF/DOCX) to auto-extract
3. Select travel purpose and enter travel details
4. Add rehabilitation summary and supporting documents list
5. Click **Generate Letter** — the letter preview appears on the right
6. Edit individual sections as needed using the tab editor
7. Click **Download DOCX** to export the formatted letter

Letters are saved to Supabase automatically and can be retrieved from the saved letters panel, filtered by case manager.

## Legal Notes

The statute lookup system is date-aware and implements the *Canada (Citizenship and Immigration) v. Tran* principle: criminality is assessed against the statute in force at the date of the offence, not the date of application. This affects DUI thresholds (pre/post 2018), cannabis offences (pre/post legalization), and other amended statutes.

Special procedural dispositions (CWOF, PBJ, ARD, expungements, deferred adjudication, diversion) are recognized and noted appropriately in the generated letter.

## Environment Variables

| Variable | Required | Description |
|---|---|---|
| `ANTHROPIC_API_KEY` | Yes | Your Anthropic API key |
| `SUPABASE_URL` | Yes | Your Supabase project URL |
| `SUPABASE_ANON_KEY` | Yes | Supabase anon/public key |
| `SUPABASE_SERVICE_ROLE_KEY` | Yes | Supabase service role key (server-side only) |
| `TRP_PORT` | No | Port for local dev (default: `8081`) |
