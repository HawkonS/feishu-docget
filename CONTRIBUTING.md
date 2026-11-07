# Contributing

Thank you for taking the time to improve feishu-docget. Contributions are welcome, including bug reports, focused fixes, tests, and documentation updates.

## Development setup

The project uses Python 3.10 in CI. Install the project dependencies with:

```bash
python -m pip install -r requirements.txt
```

Before submitting a pull request, run the unit tests and Python syntax check:

```bash
python -m unittest discover -s tests -p 'test_*.py' -v
python -m compileall ./src
```

The unit test command is also run by GitHub Actions. If you change the web interface, include a screenshot in your pull request when it helps reviewers understand the change.

## Commit messages

Use [Conventional Commits](https://www.conventionalcommits.org/) for every commit message and pull request title:

```text
type(scope): summary
```

- **English is required** for the entire title, including the scope and summary.
- **A scope is required**. Start it with a lowercase English letter and use only lowercase English letters, digits, and hyphens.
- **The entire commit message must be a single title line**. Do not add a body, footers, trailers, or a second `-m` argument.
- **The complete title must be 72 characters or fewer**, including the type, scope, spaces, and punctuation.
- Keep the type and scope lowercase. Start the summary with a lowercase imperative verb, such as `add`, `fix`, `remove`, `update`, or `document`.
- Proper names and technical identifiers in the summary may retain their original capitalization, such as `Android`, `GitHub`, `SDK`, and `JAVA_HOME`.
- Do not end the title with a period or whitespace, and do not use emoji.
- Keep each commit focused on one logical change. Describe the actual change; avoid vague summaries such as `update code` or `fix bug`.
- For an incompatible change, put `!` after the scope and before the colon: `type(scope)!: summary`.
- Put context, issue references, validation details, and migration notes in the pull request description.
- Apply the same rules to pull request titles and final squash-merge commit messages. When squash-merging, keep only the title and remove any automatically generated body or trailers.

Use one of these types:

| Type | Use for |
| --- | --- |
| `feat` | New user-visible functionality |
| `fix` | Bug, regression, or compatibility fixes |
| `refactor` | Code restructuring without behavior changes |
| `ui` | Visual styling or interaction adjustments |
| `perf` | Performance improvements |
| `docs` | Documentation changes |
| `test` | Test additions or updates |
| `build` | Build configuration, dependencies, tooling, or packaging |
| `ci` | CI workflows and automation |
| `chore` | Maintenance that fits no other type |
| `revert` | Reverting an earlier commit |

Use `feat` for new functionality and `fix` for defects, including changes to the UI. Use `ui` for other visual or interaction adjustments.

Preferred scopes:

| Scope | Area |
| --- | --- |
| `web` | User-facing pages and frontend interactions |
| `admin` | Administration features and interface |
| `api` | Backend endpoints and application services |
| `converter` | Document conversion and export |
| `auth` | OAuth login, sessions, and permissions |
| `cli` | Command-line tools and local scripts |
| `config` | Configuration loading and schema |
| `template` | Word templates and style resources |
| `build` | Build configuration, dependencies, tooling, and packaging |
| `ci` | CI workflows and automation |
| `docs` | README files and project documentation |
| `repo` | Repository-level configuration and contribution guidelines |

Prefer the scopes listed above. Use another concise scope only for a distinct area not covered by them, and keep it consistent. Tests should normally use the scope of the code they cover.

Examples:

```text
feat(converter): preserve table cell colors in exported docx
fix(admin): load contact author email into frontend
test(api): cover invalid download paths
docs(repo): document open-source contribution workflow
feat(api)!: change the download response format
```

## Pull requests

- Keep changes focused and explain the problem they solve.
- Add or update tests when behavior changes, and include the commands you ran in the pull request description.
- Update the README or other user-facing documentation when behavior or configuration changes.
- Do not include credentials, `feishu-docget.properties`, logs, exported files, databases, or private templates.
- Report security-sensitive issues privately to the maintainer instead of publishing exploit details in an issue.

## License

By contributing, you agree that your contributions are licensed under the [Apache License 2.0](./LICENSE).
