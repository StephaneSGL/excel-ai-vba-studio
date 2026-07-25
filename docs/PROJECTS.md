# Carte des projets locaux

Le dépôt `C:\Users\Stephane\Desktop\Code\excel-ai-vba-studio` est la source
canonique de l’extension.

Les paquets VSIX locaux sont générés dans `output/vsix`, jamais à la racine.

Deux prototypes locaux restent utiles :

- `office-workbench`, branche `feat/xlsm-native-editing` : édition native XLSM
  ciblée et tests de préservation.
- `tu-peux-trouver-normalement-il-y`, nom interne `Office Workbench` : pont
  natif C# / Excel COM et protocole JSON.

Ils servent de sources techniques. Ils ne doivent pas remplacer le dépôt
canonique ni être poussés vers son remote.

## Flux de travail

1. Créer une branche depuis la version canonique la plus récente.
2. Porter une capacité isolée depuis un prototype.
3. Adapter ses tests au dépôt canonique.
4. Exécuter `npm run validate`.
5. Faire relire le diff.
6. Commit, push et PR seulement après autorisation.

## Chantiers

- Étendre l’outil transactionnel d’écriture VBA livré en `0.2.0` à la création sûre de designers UserForm.
- Réintégrer le pipeline d’édition native XLSM dans la version `0.1.7+`.
- Choisir les composants utiles du pont C# sans maintenir deux extensions
  concurrentes.
