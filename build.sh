#!/bin/bash
set -e

# 1. Prendi le ultime modifiche da Windows (App.tsx ecc.)
git pull --rebase

# 2. Build
rm -f docs/assets/index-*.js docs/assets/index-*.css
npm run build

# 3. Commit docs/
git add docs/
git diff --cached --quiet && echo "Niente da committare." && exit 0
git commit -m "build"

# 4. Push con risoluzione automatica conflitti su docs/index.html
while ! git push; do
  git pull --rebase -X ours
done

echo "Deploy completato."
