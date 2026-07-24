# Contribuer / Contributing

Merci de contribuer à Excel AI & VBA Studio. Les échanges peuvent être rédigés en français ou en anglais.

## Avant de commencer

- Recherchez les [issues](https://github.com/StephaneSGL/excel-ai-vba-studio/issues) et [discussions](https://github.com/StephaneSGL/excel-ai-vba-studio/discussions) existantes.
- Utilisez une discussion pour une question générale et une issue pour un bug reproductible ou une proposition précise.
- Signalez toute vulnérabilité uniquement via le [formulaire privé GitHub](https://github.com/StephaneSGL/excel-ai-vba-studio/security/advisories/new).
- Ne partagez jamais de classeur réel, de données confidentielles, d’identifiants ou de code VBA propriétaire. Créez un exemple synthétique minimal.

## Contributions de code et propriété

Ce dépôt est public et **source-available**, mais il n’est pas distribué intégralement sous une licence open source. Les issues, rapports de bugs, tests synthétiques et propositions d’architecture sont les bienvenus.

Une contribution substantielle de code doit être discutée avec le mainteneur avant l’ouverture d’une pull request. Avant fusion, le mainteneur peut demander un accord de contribution ou de cession distinct afin de préserver sa capacité à maintenir et licencier le projet. L’ouverture d’une pull request ne transfère pas automatiquement la propriété du code.

En soumettant du contenu, vous garantissez que vous avez le droit de le proposer, qu’il ne contient pas de code confidentiel ou incompatible, et vous acceptez qu’aucune fusion ne soit acquise. Ne soumettez pas de code substantiel si vous n’acceptez pas le régime décrit dans [LICENSE](LICENSE) et [LICENSING.md](LICENSING.md).

Les mentions de licence et d’attribution d’Office Viewer, de x-data-spreadsheet, de Vditor et des autres composants tiers doivent être conservées.

## Développement local

Prérequis : Node.js 22, npm, Visual Studio Code 1.95 ou ultérieur. Windows x64 et Microsoft Excel desktop sont nécessaires pour tester l’intégration native Excel, COM et VBA.

```powershell
git clone https://github.com/StephaneSGL/excel-ai-vba-studio.git
cd excel-ai-vba-studio
npm ci
npm run validate
```

Pour produire un VSIX local :

```powershell
npm run package
```

Les chemins d’export, l’automatisation Excel, les macros et les formats `.xlsm`, `.xls` et `.xlsb` sont sensibles. Toute modification de ces zones doit préserver les garde-fous de sécurité, ne jamais exécuter de macro pendant l’analyse et utiliser uniquement des fichiers de test synthétiques.

## Pull requests

- Gardez chaque pull request ciblée et expliquez le problème résolu.
- Décrivez les tests manuels et automatiques effectués.
- Mettez à jour la documentation et `CHANGELOG.md` lorsque le comportement visible change.
- Préservez toutes les mentions de propriété, de licence et d’attribution.
- N’ajoutez aucune dépendance ou source copiée sans documenter sa provenance et sa licence.

---

Contributions are welcome in French or English. This repository is public and source-available, but it is not wholly open source. Discuss substantial code contributions with the maintainer before opening a pull request; a separate contributor or assignment agreement may be required before merge. Use synthetic workbooks only, run `npm run validate`, preserve all third-party notices, and report vulnerabilities privately.
