# Git Quick Reference

A simple guide to the four core Git actions for day-to-day use in VS Code.

---

## The Everyday Flow

```
Pull → Make changes → Stage → Commit → Push
```

Always **Pull** before you start. Always **Push** when you're done.

---

## The Four Core Actions

### Pull
Downloads the latest version of the repo from GitHub to your local machine.
Do this first every time you sit down to work, especially after using another machine.

**VS Code:** `Ctrl+Shift+P` → **Git: Pull**
**Terminal:**
```bash
git pull
```

---

### Stage
Selects which changed files to include in your next commit.
You may have changed multiple files but only want to save some of them right now.

**VS Code:** Source Control panel (`Ctrl+Shift+G`) → click **+** next to each file
**Terminal:**
```bash
git add filename.ps1        # Stage a specific file
git add .                   # Stage all changed files
```

---

### Commit
Creates a permanent local snapshot of your staged files with a description of what changed.
Nothing goes to GitHub yet — this is just a local save point.

**VS Code:** Type a message in the box at the top of Source Control → click **✓ Commit**
**Terminal:**
```bash
git commit -m "Your message here"
```

> **Write useful commit messages.** `Fix batch size bug in classifier` is far more helpful than `changes` when you're looking back through history.

---

### Push
Uploads your local commits to GitHub.
Only after a push is your work visible in the browser and available on your other machines.

**VS Code:** Click **Sync Changes** in the Source Control panel
**Terminal:**
```bash
git push
```

---

## Common Scenarios

### Starting a session
```bash
git pull
```

### Saving and publishing your work
```bash
git add .
git commit -m "Describe what you changed"
git push
```

### Checking what has changed locally
```bash
git status
```

### Viewing recent commit history
```bash
git log --oneline
```

---

## VS Code Keyboard Shortcuts

| Action | Shortcut |
|---|---|
| Open Source Control panel | `Ctrl+Shift+G` |
| Command palette | `Ctrl+Shift+P` |
| Open terminal | `Ctrl+`` ` `` |

> **Mac users:** Replace `Ctrl` with `Cmd` for all shortcuts above.

---

## Cloning a Repo to a New Machine

1. Open VS Code
2. `Ctrl+Shift+P` (or `Cmd+Shift+P` on Mac) → **Git: Clone**
3. Paste the repo URL: `https://github.com/RudeOne-sys/powershell-scripts.git`
4. Choose a local folder
5. Open the cloned folder when prompted

---

## First-Time Git Setup (new machine)

Run these once in the terminal before your first commit:

```bash
git config --global user.name "Your Name"
git config --global user.email "your@email.com"
```

Use the same email as your GitHub account.
