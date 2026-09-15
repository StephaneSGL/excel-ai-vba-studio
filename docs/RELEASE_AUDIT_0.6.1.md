# Audit et recette — candidate 0.6.1

Date : 15 septembre 2026. Base auditée : `22f44a28c2cde33f2bd7b6a849ef69124359e37d` (`main`, version 0.6.0).

## Décision de livraison

La candidate reste **Preview**. Les contrôles locaux automatisés permettent de préparer un VSIX, mais ne prouvent pas que chaque bouton et chaque intégration fonctionne dans toutes les versions d’Office. Excel desktop n’est pas installé sur le poste d’audit. La recette COM/Excel et la recette utilisateur sur le second PC restent obligatoires avant une déclaration de version finale.

`npm run validate:native-release` doit réussir sur un poste Windows/Excel approprié. Il échoue volontairement si un prérequis natif manque ; `npm run validate` annonce explicitement les tests natifs ignorés. Aucun réglage AccessVBOM, ActiveX, registre ou sécurité Office n’a été modifié pour contourner ce manque.

## Correctifs livrés

| Constat | Correction et garde-fou |
| --- | --- |
| Quatre `pattern` avec `\p{L}`/`\p{N}` étaient encore présents dans le manifeste de la base auditée. | Retrait de ces seuls motifs, conservation des longueurs et de la validation métier Unicode avec drapeau `u`. Compilation des schémas et contrôle récursif des motifs portables ajoutés. |
| Deux outils VBA pouvaient résoudre implicitement le classeur actif au moment de l’écriture. | Chemin absolu obligatoire, cible canonique complète dans la confirmation, aucun repli sur l’éditeur actif. Résolution locale partagée pour les outils mutateurs ; tests du changement de classeur entre préparation et invocation. |
| Remplacer utilisait un objet feuille comme indice de tableau. | Indexation corrigée ; remplacement littéral de `$&`, métacaractères et recherches multi-feuilles testés ; cellules verrouillées préservées. |
| Le contrôle TypeScript principal ne couvrait pas la webview. | Deux contrôles distincts, inclus dans `validate` ; types ExcelJS, images, protection et données de feuilles corrigés, messages d’ouverture validés. |
| Types VS Code plus récents que la version minimale annoncée. | `@types/vscode` fixé à 1.95.0 ; `@types/node` fixé à 20.19.43 pour le moteur ciblé. Démarrage réel testé sur VS Code 1.95.0 et 1.137.0. |
| Quatre dépendances signalées par npm audit. | Versions verrouillées : fast-uri 3.1.8, js-yaml 4.3.2, nanoid 3.3.19, qs 6.16.0. Audit npm : aucun avis connu remonté lors du contrôle, ce qui ne constitue pas une garantie d’absence de vulnérabilité. |
| Ancien code de visionneuses et historique Git sans chemin d’entrée produit. | Suppression de 100 fichiers / 23 459 lignes. Le graphe du produit conserve ses 167 fichiers source atteignables avant/après ce nettoyage. Suppressions récupérables dans Git. |
| Définitions d’images et assainissement d’URL d’icônes dupliqués. | Types d’images partagés et fonction d’assainissement unique ; tests existants conservés. |
| La préparation de release écrivait le manifeste avant de vérifier le lockfile. | Vérification des deux entrées avant modification ; test de refus sans modification du manifeste. |
| XLAM était annoncé par l’outil alors que la validation des chemins le refusait déjà. | Promesse retirée du manifeste et documentation clarifiée ; pas d’ajout de support natif non testé. |

## Périmètre des vérifications

| Domaine | Preuve / portée | Limite |
| --- | --- | --- |
| Installation logique et activation VS Code | Instance Extension Host isolée : activation, 12 commandes, 4 outils propres à l’extension, ouverture de l’éditeur CSV ; versions 1.95.0 et 1.137.0. | Ce smoke test ne certifie pas le rendu visuel de chaque panneau. |
| Outils IA | Schémas compilés, validations d’opérations, cibles explicites et confirmations testées. | Les tests de ciblage isolent les frontières fichiers/COM ; pas de conversation Copilot connectée testée de bout en bout. |
| Classeur et grille | Tests de lecture du classeur synthétique à cinq feuilles, recherche/remplacement, tables disjointes, noms Unicode NFC, images et arrière-plans après sérialisation OOXML. | Les tests des modèles ne remplacent pas une recette de tous les gestes UI. |
| Tableaux et graphiques | Catalogues, concepteur, parsing SERIES, plages, métadonnées OOXML et génération de transactions vérifiés. | Création/modification réelle dans Excel à tester sur le poste Office. |
| Écriture VBA hors COM | Helper Windows et scénarios de modules existants/nouveaux, code UserForm, refus d’écriture périmée ; tests Python du remplacement atomique et du rollback. | Le designer natif et l’inventaire Excel nécessitent la recette séparée. |
| Sécurité | Signatures de package, macros XLM, ZIP, limites de lecture, transport, flux NTFS et refus pré-COM ; vérifications statiques complémentaires. | Les décisions effectives de politique Office et les transactions COM positives ne sont pas démontrées sur ce poste. |
| Build et distribution | Build de production, contrôle TypeScript hôte/webview, budget de chargement initial, liste autorisée du VSIX et licences. | Un VSIX local n’est pas une publication Marketplace ni l’artefact reconstruit d’un tag CI. |

## Ce qui n’a pas été supprimé aveuglément

`npm run audit:source` produit une liste de candidats non atteints depuis les entrées de build. Ce n’est pas une liste automatique de fichiers à supprimer : imports de types, shims injectés par Vite, assets et modules utilisés par des tests demandent une revue séparée. Les composants HTTP/archives encore couverts par les régressions de sécurité et les attributions de licences ont été conservés. Ils restent exclus du VSIX par sa liste autorisée.

Le JavaScript tiers embarqué n’est pas intégralement contrôlé par TypeScript. Une campagne de tests visuels et de couverture des interactions est encore utile avant un statut stable ; aucune affirmation de couverture à 100 % n’est faite.

## Recette restante avant promotion finale

1. Sur un poste de test Windows x64 avec Excel desktop, cloner cette révision, exécuter `npm ci`, `npm run validate`, puis `npm run validate:native-release`. Conserver les logs et les versions exactes d’Office/Windows. Utiliser exclusivement les fixtures synthétiques et les permissions déjà approuvées par l’utilisateur.
2. Installer le VSIX candidat dans un profil VS Code de test sur le second PC ; confirmer la disparition de l’erreur du manifeste. Le message brut et la version exacte qui produisaient l’erreur initiale n’ont pas été fournis.
3. Recette visuelle : CSV/TSV/XLSX/XLSM, navigation, formats/formules, undo/redo, recherche et remplacement, images, tableaux, graphiques, sauvegarde, rechargement et récupération après fermeture.
4. Recette native : conversion XLSX vers copie XLSM sans modifier la source, modules VBA, UserForms/contrôles/événements, boutons, ActiveX autorisé et refusé, rollback, sauvegardes, classeur actif changé et modifications externes concurrentes.
5. Recette de refus : fichier signé/protégé/verrouillé, Protected View, accès VBA refusé, chemin distant/reparse point, limite de taille et annulation. Les politiques de sécurité ne doivent jamais être abaissées automatiquement.
6. Après revue et CI verte, produire le tag et utiliser l’artefact reconstruit par le workflow de release. Le runbook Marketplace reste une étape distincte ; aucune publication Marketplace automatique n’est configurée.

Méthode de smoke test : [documentation officielle VS Code — Testing Extensions](https://code.visualstudio.com/api/working-with-extensions/testing-extension).
