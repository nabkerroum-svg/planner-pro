# 🚀 Déploiement Planner Pro

## Option 1 — Vercel (recommandé, 3 commandes)

```bash
npm install
npm install -g vercel
vercel
```
→ Ton site sera en ligne sur `https://planner-pro-xxx.vercel.app`

---

## Option 2 — Netlify (glisser-déposer)

```bash
npm install
npm run build
```
Puis va sur https://netlify.com/drop et glisse le dossier `build/`

---

## Option 3 — GitHub Pages

1. Crée un repo GitHub nommé `planner-pro`
2. Dans package.json, remplace `TON_PSEUDO` :
   `"homepage": "https://TON_PSEUDO.github.io/planner-pro"`

```bash
npm install
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/TON_PSEUDO/planner-pro.git
git push -u origin main
npm run deploy
```
→ Site en ligne sur `https://TON_PSEUDO.github.io/planner-pro`

---

## Tester en local avant de publier

```bash
npm install
npm start
```
→ Ouvre http://localhost:3000
