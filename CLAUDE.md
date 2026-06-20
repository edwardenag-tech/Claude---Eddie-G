# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Single-file browser game — no build step, no dependencies, no package manager. Everything lives in `tic-tac-toe.html`.

## Running the Project

Open `tic-tac-toe.html` directly in a browser. No server required.

## Architecture

All code is self-contained in one HTML file with three sections:

- **CSS** (`<style>` block) — layout and theming via CSS custom properties (`--bg`, `--cell`, `--primary`, etc.)
- **HTML** (`<body>`) — static 3×3 grid of `.cell` divs, a status line, score counters, and a reset button
- **JavaScript** (`<script>` block) — all game logic; no frameworks or external libraries

### Key JS structures

- `WINS` — array of the 8 winning index triplets
- `THEMES` — array of 8 color theme objects; a new theme is applied after each win or draw via `nextTheme()`
- `init()` — resets board state; only initializes `scores` and `themeIndex` on first call
- `checkWin()` — iterates `WINS` and returns the winning triplet or `null`
- Event listeners on `#board` (click delegation) and `#reset`

## Git Workflow

After every change: commit with a clean descriptive message and push to GitHub immediately. The remote is `https://github.com/edwardenag-tech/Claude---Eddie-G.git`.
