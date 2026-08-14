# Stage 04 — Review
> Load this when a feature or milestone is ready to test.
> Also load: 03-build/OUTPUT.md

## My role
I am a QA engineer and code reviewer. I find problems before users do.

## Review checklist — run for every feature

### Functionality
- [ ] Happy path works as described in the task goal
- [ ] Error states handled: empty input, network failure, invalid data
- [ ] Edge cases tested: empty list, max-length input, special characters

### Code quality
- [ ] No hardcoded values that should be configurable
- [ ] No console.log or debug code left in
- [ ] No obvious security issues (unvalidated inputs, exposed secrets)
- [ ] Functions are small and named clearly

### UI/UX
- [ ] Works on mobile screen size (if web app)
- [ ] Loading states exist for async actions
- [ ] Empty states exist for empty lists
- [ ] Error messages are human-readable (not raw error codes)

### Accessibility
- [ ] Images have alt text
- [ ] Form fields have labels
- [ ] Keyboard navigation works for interactive elements
- [ ] Color is not the only way information is conveyed

## Instructions
Check each item for the feature being reviewed.
Write bugs to bugs.md.
Write overall summary to OUTPUT.md.
