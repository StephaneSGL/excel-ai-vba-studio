# Excel AI & VBA Studio — livraison 0.6.2 Preview

Vérifié le 21 septembre 2026. **Paquet installable pour essais ; pas une certification de toutes les fonctions Excel/VBA.**

## Installer et commencer

1. Utiliser Windows x64 et VS Code 1.95 ou ultérieur.
2. Télécharger le VSIX Windows x64 joint à la [release v0.6.2](https://github.com/StephaneSGL/excel-ai-vba-studio/releases/tag/v0.6.2), pas le ZIP des sources. Dans VS Code : **Extensions → … → Installer à partir d’un VSIX**, choisir le fichier, puis recharger la fenêtre.
3. Appuyer sur **Ctrl+Maj+P**, puis sélectionner **Excel AI & VBA Studio : Accueil et diagnostic**. Vérifier que l’accueil affiche **0.6.2**.
4. Cliquer sur **Ouvrir un classeur…** et commencer avec une copie de test XLSX, CSV ou TSV. La commande ouvre directement la grille intégrée, sans changer les associations d’éditeurs.
5. Pour une anomalie, utiliser **Copier le diagnostic** et conserver le message d’erreur exact. Ne pas envoyer de classeur confidentiel.

L’installation du VSIX ne demande pas Node.js, npm ou Python. Les tests de livraison utilisent des profils VS Code temporaires : **le profil habituel de l’utilisateur n’a pas été modifié**.

## Ce qui a été amélioré

- Accueil en français, avec thèmes clair/sombre/contraste élevé, boutons accessibles au clavier et disposition adaptée aux panneaux étroits.
- Actions d’ouverture et d’accueil dans l’explorateur, même lorsqu’aucun classeur n’est ouvert. L’ancien message sans action n’est plus le seul point de départ.
- Diagnostic local : Windows x64, espace de travail approuvé, PowerShell, enregistrement Excel dans les deux vues du registre, composant VBA embarqué et présence des quatre outils IA.
- Explications séparant la grille sans Excel, les fonctions natives avec Excel et le chat IA facultatif. Un enregistrement Excel n’est jamais présenté comme une validation COM ou VBA.
- Rapport sans chemin de fichier, contenu de classeur, identifiant de compte ou clé. Aucun envoi réseau par le diagnostic ; aucune modification des réglages Office.
- Ouverture bornée aux formats pris en charge, contrôle de confiance, annulation du sélecteur et liste fermée des actions reçues depuis l’accueil. Aucun chemin ou nom de commande fourni par la page n’est exécuté.

Les correctifs 0.6.1 du manifeste (`\p{L}` / `\p{N}`), du ciblage explicite des classeurs, du collage et des validations métier restent présents.

## Résultats vérifiés

| Contrôle | Résultat et portée |
| --- | --- |
| Dépendances | Réinstallation par `npm ci` depuis le lockfile ; l’installation locale précédente avait dérivé. Pas de mise à jour des versions verrouillées. |
| `npm run validate` | Réussite : manifeste, garde-fous, régressions, TypeScript hôte/webview, build, licences, performances et contenu du paquet. Cinq suites ou parties natives signalées comme ignorées faute d’Excel. |
| `npm audit --json` | Aucun avis de vulnérabilité connu remonté lors du contrôle. Ce n’est pas une garantie d’absence de vulnérabilité. |
| VSIX sur VS Code 1.95.0 | Installation effective dans un profil isolé, chargement vérifié depuis ce paquet, activation, 14 commandes, 4 outils de l’extension, accueil/diagnostic et ouverture guidée d’un CSV : réussite. |
| VSIX sur VS Code stable 1.138.0 | Mêmes contrôles : réussite. Le runner prend en charge le nouvel emplacement versionné du CLI de VS Code. |
| Accueil | Rendu inspecté dans le navigateur sur un état synthétique : sombre, clair, panneau de 600 px sans débordement horizontal, aide et message de bouton. Les commandes réelles sont couvertes séparément par les tests hôte/unitaires. |
| Paquet | 45 entrées dans le VSIX, dont 43 fichiers produit ; 10 122 165 octets. |
| `npm run validate:native-release` | Échec explicite attendu : Excel COM ou l’accès VBA nécessaire n’est pas disponible sur ce poste. |

Les tests hôte ne certifient pas toutes les interactions de la grille, les conversions complexes, les politiques Office, les UserForms exécutés dans Excel ni une conversation Copilot connectée. Les journaux du VS Code isolé signalent notamment l’absence de compte IA ; aucune connexion personnelle n’a été utilisée.

Le build affiche un avertissement de compatibilité future du chargeur de configuration Vite ; le build actuel réussit. Il ne s’agit pas d’une erreur de chargement du VSIX.

## Intégrité de la candidate locale testée

SHA-256 du VSIX :

```text
829df0a50458fa13be42fe81cacce472a7aaca1fda3d3fad7028bde22f4058b6
```

Ce digest désigne le build local décrit ci-dessus, pas nécessairement l’artefact reconstruit en CI. Pour le téléchargement public, utiliser le fichier **SHA256SUMS.txt joint à la release** : son VSIX provient du workflow du tag et de son helper reconstruit. Le paquet local 0.6.1 précédent est conservé.

## Dernière étape sur le PC équipé d’Excel

Sur des fichiers synthétiques ou des copies jetables : vérifier ouverture, modification, sauvegarde, réouverture, conservation des formules/objets, opérations VBA et récupération après erreur. Puis exécuter `npm run validate:native-release` depuis les sources et conserver le journal avec les versions de Windows, Office et VS Code.

Ne pas activer automatiquement AccessVBOM, les macros, les emplacements approuvés ou ActiveX. Toute permission requise doit correspondre aux choix de l’utilisateur et à la politique de son organisation.

## État Git et publication

La version **0.6.2 Preview** est préparée depuis `codex/release-0.6.1-hardening`, à partir de `756baeb`. Sa publication GitHub utilise le tag `v0.6.2`, sans fusion automatique dans `main`, publication Marketplace ni remplacement du brouillon GitHub 0.6.1. Le statut Preview et les prérequis Excel restent inchangés.

La configuration CI teste l’installation du VSIX sur 1.95.0 et stable. Le workflow du tag reconstruit et vérifie le helper natif, puis produit le VSIX destiné à la release. Les résultats ci-dessus décrivent la recette locale initiale ; les résultats distants et le digest de l’artefact publié sont consignés dans les notes de release. Aucune réussite CI ne remplace la recette Excel/COM.
