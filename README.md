# Princess Parts Scraper

1) "C:\perso\00 formation Data analyst\Data Management\1 Applications perso\11 ML nautiC\pieces\Prompt pour refaire le scrapping.docx"

2)A faire après validation du nouveau fichier excel, 

«  implemente moi les données avec le script d’import import-xlsx-to-sqlite.ps1  « 

Fichiers créés
Base SQLite réelle : princess_parts_analytics.sqlite.db
Script d’import manuel Excel vers SQLite : import-xlsx-to-sqlite.ps1
Générateur SQL analytique utilisé par l’import : generate-sqlite-analytics-sql.ps1

Ce que fait le script d’import import-xlsx-to-sqlite.ps1

•	ouvre un fichier Excel
•	lit la première feuille par défaut, ou une feuille précise si tu la donnes
•	reconstruit un CSV propre
•	génère le SQL analytique SQLite
•	l’importe dans la base .db
•	met à jour les données de façon idempotente sur le couple date + url
•	nettoie les runs orphelins éventuels


3) Mise a jour GitHub
Apres un nouveau scraping, tu peux pousser les changements vers GitHub avec une seule commande PowerShell :

```powershell

pSet-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass; .\scripts\update-github.ps1 -CommitMessage "Mise a jour apres nouveau scraping"

Le script :
- verifie que le dossier est bien un depot Git
- detecte s'il y a des changements
- ajoute tous les fichiers modifies
- cree un commit
- pousse sur la branche main du depot GitHub




Ce projet contient un Cloudflare Worker qui scrappe https://parts.princess.co.uk/ via fetch() et HTMLRewriter pour extraire les produits et les exporter en CSV.

## Champs extraits

- category
- name
- description
- specifications
- model_number
- price_gbp_numeric
- stock
- delivery_delay
- url

## Pourquoi le scraping est batché

Le sitemap produit expose 271 URLs. Un scraping complet en une seule requete depasserait facilement les limites de sous-requetes d'un Cloudflare Worker. Le Worker expose donc des endpoints batches.

## Endpoints

- /manifest
- /scrape.json?offset=0&limit=20
- /scrape.csv?offset=0&limit=20

## Utilisation

1. Installer Node.js.
2. Installer les dependances avec npm install.
3. Lancer le Worker en local avec npm run dev ou le deployer avec npm run deploy.
4. Consolider les lots CSV avec scripts/export-worker-batches.ps1.

Exemple PowerShell:

```powershell
.\scripts\export-worker-batches.ps1 -BaseUrl https://princess-parts-scraper.<votre-sous-domaine>.workers.dev -OutputFile .\princess_parts_products.csv
```

## Notes de parsing

- Les URLs produit sont lues depuis product-sitemap.xml.
- Les categories sont lues dans .product_meta .posted_in.
- Le Model Number est lu dans .product_meta .sku_wrapper .sku.
- Le descriptif est lu depuis les paragraphes de .woocommerce-product-details__short-description.
- Les Specifications et le Delivery delay sont lus dans l'accordeon #accordion.
- price_gbp_numeric est normalise en decimal avec point et deux decimales, sans separateur de milliers.
- Le stock est normalise en valeur numerique: quantite si presente, sinon 0.

