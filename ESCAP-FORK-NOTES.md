# ESCAP Fork Notes

This is the **ESCAP fork** of `@turbodocx/html-to-docx`, used by the
[ESCAP Document Center](https://unescap.visualstudio.com/ESCAP-Document-Center/)
for its lossless DOCX↔HTML bridge feature.

## Why this fork exists

We added a single named export, `htmlToOoxmlFragment(html)`, that returns
the OOXML body fragment for a given HTML string without packaging it as a
full `.docx` file. The Document Center's `HtmlModule` for docxtemplater
needs the fragment, not a complete document. Everything else is upstream.

## Mirrors

The fork lives at **two** Git remotes simultaneously, both fully writable:

| Remote   | URL                                                                          | Role                       |
|----------|------------------------------------------------------------------------------|----------------------------|
| `origin` | `https://github.com/escap-imcts-dtu/html-to-docx.git`                        | Primary; consumed by ESCAP |
| `ado`    | `https://unescap.visualstudio.com/ESCAP-Document-Center/_git/html-to-docx`   | Backup mirror              |

If one remote is ever deleted by mistake, the other holds the full history.

## How ESCAP installs us

ESCAP's `package.json` pins to a **specific commit SHA**:

```json
"@turbodocx/html-to-docx": "github:escap-imcts-dtu/html-to-docx#<sha>"
```

We commit `dist/` to this repo so consumer installs are zero-build (no
rollup runs at install time). When you change source files:

```bash
npm run build
git add dist/ && git commit
./scripts/push-mirrors.sh main
```

Then bump the SHA in ESCAP's `package.json`.

## After cloning fresh

```bash
./scripts/setup-mirrors.sh   # configures origin to push to both URLs
./scripts/push-mirrors.sh    # one-shot push to both (uses az login for ADO)
```

## Pulling upstream updates

```bash
git fetch upstream
git merge upstream/main      # resolve any conflicts in our changes
npm run build
git add dist/ && git commit
./scripts/push-mirrors.sh main
```
