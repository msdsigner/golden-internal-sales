# Skill: Naming Conventions
> Load this when creating files, functions, components, or database fields.
> Consistent naming reduces the cognitive load of navigating a codebase.

## Files
- React components: PascalCase.jsx  (e.g. UserProfile.jsx)
- Utility functions: camelCase.js   (e.g. formatDate.js)
- Pages (Next.js): kebab-case.jsx   (e.g. user-settings.jsx)
- API routes: kebab-case            (e.g. /api/user-profile)

## Code
- Variables and functions: camelCase         (e.g. getUserById)
- Classes and components: PascalCase         (e.g. UserService)
- Constants: SCREAMING_SNAKE_CASE            (e.g. MAX_RETRIES)
- Boolean variables: starts with is/has/can  (e.g. isLoading, hasError)

## Database
- Tables: snake_case, plural   (e.g. user_profiles, blog_posts)
- Columns: snake_case          (e.g. created_at, user_id)
- Foreign keys: [table]_id     (e.g. user_id, post_id)

## Git
- Branches: type/short-description  (e.g. feat/user-auth, fix/login-redirect)
- Commits: type: description        (e.g. feat: add login form)
  Types: feat | fix | chore | docs | refactor | test

## Environment variables
- SCREAMING_SNAKE_CASE, prefixed by app name  (e.g. MYAPP_DATABASE_URL)
- Never commit .env files
