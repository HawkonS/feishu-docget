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

Use this format for every commit:

```text
type(scope): description
```

- Write the entire title in English.
- A scope is required. Use a concise, lowercase name for the affected area, such as `web`, `admin`, `api`, `converter`, `auth`, `cli`, `config`, `template`, `docs`, or `repo`.
- Use one title line only. Do not add a commit body or additional `-m` argument.
- Start the description with a lowercase imperative verb, such as `add`, `fix`, `restore`, `update`, or `document`.
- Keep the title concise, preferably no more than 72 characters. Do not end it with a period or use emoji.

Common types:

| Type | Use for |
| --- | --- |
| `feat` | New functionality |
| `fix` | Bug or compatibility fixes |
| `refactor` | Code restructuring without changing behavior |
| `perf` | Performance improvements |
| `docs` | Documentation changes |
| `test` | Test additions or updates |
| `build` | Build, dependency, or packaging changes |
| `ci` | CI workflows and automation |
| `chore` | Maintenance that does not fit another type |

Examples:

```text
feat(converter): preserve table cell colors in exported docx
fix(admin): load contact author email into frontend
test(api): cover invalid download paths
docs(repo): document open-source contribution workflow
```

Keep each commit focused on one logical change. Use the same format for pull request titles when the changes will be squash-merged.

## Pull requests

- Keep changes focused and explain the problem they solve.
- Add or update tests when behavior changes, and include the commands you ran in the pull request description.
- Update the README or other user-facing documentation when behavior or configuration changes.
- Do not include credentials, `feishu-docget.properties`, logs, exported files, databases, or private templates.
- Report security-sensitive issues privately to the maintainer instead of publishing exploit details in an issue.

## License

By contributing, you agree that your contributions are licensed under the [Apache License 2.0](./LICENSE).
