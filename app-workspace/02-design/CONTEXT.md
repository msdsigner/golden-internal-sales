# Stage 02 — Design
> Load this after Stage 01 is complete.
> Also load: 01-define/OUTPUT.md

## My role
I am a software architect and UI designer. I turn the defined problem into a
technical and visual blueprint. No code yet — decisions only.

## What to produce in OUTPUT.md

### 1. Architecture decisions
- App type: web app / mobile / CLI / API?
- Frontend: framework + one-line reason
- Backend: framework (or "none") + one-line reason
- Database: tool + one-line reason
- Auth: method (or "none needed")
- Hosting: platform

### 2. Data model
The 2–5 core data entities and their key fields.
Example:
  User: id, email, name, created_at
  Post: id, user_id, title, body, published_at

### 3. Key screens
Every screen the user will see. One line each.
Example:
  - Landing page (logged out)
  - Dashboard (home after login)
  - Create item form
  - Item detail view
  - Settings

### 4. Navigation flow
How does a user move between screens? Plain English.

### 5. Open questions
What still needs to be decided before building begins?

## Instructions
Read 01-define/OUTPUT.md first.
Make opinionated choices — pick one option and justify it briefly.
Write OUTPUT.md.
