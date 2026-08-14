# Skill: Prompt Templates
> Copy and adapt these when starting sessions in Claude, Cursor, or other AI tools.
> Always load CLAUDE.md first — it is the DNS of the system.

---

## Start a new session (any stage)
Load CLAUDE.md, then 00-system/CONTEXT.md.
Summarize the project and ask what stage we are working on today.

---

## Run the Define stage
Load CLAUDE.md, then 01-define/CONTEXT.md.
Ask me the questions needed to fill out 01-define/OUTPUT.md.
Once you have my answers, write the file.

---

## Run the Design stage
Load CLAUDE.md, then 02-design/CONTEXT.md, then 01-define/OUTPUT.md.
Produce 02-design/OUTPUT.md.

---

## Start a build session
Load CLAUDE.md, then 03-build/CONTEXT.md, then 02-design/OUTPUT.md, then 03-build/ACTIVE-TASK.md.
Confirm you understand the active task. Then begin.

---

## Write a feature spec
Load _skills/write-feature-spec.md.
Write a spec for: [describe the feature in one sentence].
Then add the task to 03-build/TASK-LIST.md.

---

## Review code
Load CLAUDE.md, then 04-review/CONTEXT.md, then _skills/code-review.md.
Review this code and write findings to 04-review/bugs.md:
[paste code here]

---

## Debug a bug
I have a bug: [describe it]
Relevant code: [paste code]
Error message: [paste error]
Fix it. Then add an entry to 04-review/bugs.md and mark it fixed.

---

## Prepare to launch
Load CLAUDE.md, then 05-ship/CONTEXT.md, then 04-review/OUTPUT.md, then 01-define/OUTPUT.md.
Walk me through the pre-launch checklist.
