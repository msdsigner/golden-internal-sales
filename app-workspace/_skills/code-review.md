# Skill: Code Review
> Load this in Stage 04 when reviewing code.

When reviewing, always check:

1. **Does it work?** Run through the happy path mentally step by step.
2. **Does it fail gracefully?** What happens when inputs are wrong, the network is down, or the database is empty?
3. **Is it readable?** Would a stranger understand this in 6 months without asking questions?
4. **Is it safe?** Check for SQL injection, XSS, unvalidated inputs, and exposed secrets.
5. **Is it simple?** Is there a less complex way to achieve the same thing?

Format feedback as:
- OK [what is good — name it specifically]
- SUGGEST [improvement idea — not blocking, explain the tradeoff]
- FIX [must fix before shipping — explain the risk]
