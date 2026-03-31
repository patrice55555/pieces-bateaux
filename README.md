# Princess Parts Scraper

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

## Mise a jour GitHub

Apres un nouveau scraping, tu peux pousser les changements vers GitHub avec une seule commande PowerShell :

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\update-github.ps1
```

Tu peux aussi definir ton propre message de commit :

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\update-github.ps1 -CommitMessage "Mise a jour apres nouveau scraping"
```

Le script :

- verifie que le dossier est bien un depot Git
- detecte s'il y a des changements
- ajoute tous les fichiers modifies
- cree un commit
- pousse sur la branche main du depot GitHub