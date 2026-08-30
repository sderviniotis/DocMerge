# Publishing DocMerge to GitHub

Two routes. Pick one.

---

## Route A: GitHub CLI (fastest, about 3 minutes)

Install the CLI once, if you do not already have it:

```bash
brew install gh
gh auth login
```

Then from inside the `docmerge-repo` folder:

```bash
cd ~/Downloads/docmerge-repo
git init
git add .
git commit -m "DocMerge 3.1: merge txt, doc, docx and pdf into one Word document"
gh repo create docmerge --public --source=. --push \
  --description "Combine .txt, .doc, .docx and .pdf into one editable Word document. Free, open source, runs entirely on your Mac."
```

That creates the repository, pushes the code, and sets the description in one step.

---

## Route B: Web interface

1. Go to https://github.com/new
2. Repository name: `docmerge`
3. Description: *Combine .txt, .doc, .docx and .pdf into one editable Word document. Free, open source, runs entirely on your Mac.*
4. Public. **Do not** tick "Add a README", "Add .gitignore" or "Choose a licence"; this folder already has all three
5. Create repository, then run:

```bash
cd ~/Downloads/docmerge-repo
git init
git add .
git commit -m "DocMerge 3.1: merge txt, doc, docx and pdf into one Word document"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/docmerge.git
git push -u origin main
```

---

## After the first push

**1. Fix the clone URL in the README.** It currently says `YOUR_USERNAME`:

```bash
sed -i '' 's|YOUR_USERNAME|your-actual-github-username|g' README.md
git commit -am "Fix clone URL" && git push
```

**2. Add topics** so people can find it. On the repository page, click the gear icon beside "About" and add:

`macos` `python` `pdf` `docx` `document-merge` `tkinter` `offline` `privacy` `word` `productivity`

**3. Tag a release.** This gives people something to download without cloning:

```bash
git tag -a v3.1 -m "DocMerge 3.1"
git push origin v3.1
```

Then on GitHub go to Releases, draft a release from the `v3.1` tag, and paste in a short summary of what it does.

**4. Check the screenshot renders.** Open the repository page and confirm the image at the top of the README loads. If not, the path in `docs/screenshot.png` did not commit; run `git add -f docs/screenshot.png` and push again.

---

## A note on the `.gitignore`

The ignore file excludes `*.docx` to keep test output out of the repository. If you later add sample documents you *want* tracked, put them in `docs/` (already whitelisted) or force-add them with `git add -f`.

---

## Optional: attracting users

If you want the project to get traction rather than just exist:

- Post to r/macapps and r/opensource. Lead with the specific problem, not the tool
- Submit to [Awesome Mac](https://github.com/jaywcjlove/awesome-mac) via pull request
- Show HN on a Tuesday or Wednesday morning US time, titled around the problem
- Write a short blog or LinkedIn post about the build; the Gatekeeper and macOS
  Tk button issues are genuinely useful war stories for other developers

Expect the first question to be "why not just use Acrobat?". The README answers
it, but be ready to answer it again in comments.
