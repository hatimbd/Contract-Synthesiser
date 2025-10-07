### Lancer le projet `contract_synthetiser` :

Ce projet prend un fichier **Word** en entrée, détecte les modifications selon un **code couleur** (🔴 rouge pour ajouter/modifier, 🌸 rose pour supprimer), et met à jour un fichier **Excel** en sortie, en créant une nouvelle version dans une feuille dédiée.

### Prérequis :

- Python 3.8+
- Git
- pip

### Étapes d'installation :

#### 1. 📥 Cloner le dépôt GitHub :

```bash
git clone https://github.com/hatimbd/contract_synthetiser.git
cd contract_synthetiser
```

#### 2. 🐍 Créer un environnement virtuel :
```bash
python -m venv myvenv
```
#### 3. 🔛 Activer l’environnement virtuel :
**Windows** :
```bash
myvenv\Scripts\activate
```
or :  

```bash
myvenv\Scripts\activate
```
**macOS/Linux** :
```bash
source venv/bin/activate
```
#### 4. 📚 Installer les dépendances :
```bash
pip install -r requirements.txt
```

### 📄 Fichiers d'entrée/sortie :
**Entrée** : le fichier Word *input/contrat_parametrage.docx* contenant une table avec des cellules colorées :

- Rouge : ajouter ou modifier la cellule

- Rose : supprimer la cellule

**Sortie** : le fichier Excel *output/parametres_mis_a_jour.xlsx* mis à jour :

- Une nouvelle feuille Vxx est crée pour chaque version

- Les modifications sont appliquées selon le code couleur

### ▶️ Lancer le script principal :

> [!CAUTION]
>
> Avant de lancer le script `main.py`, assurez-vous que les fichiers Word (`contrat_parametrage.docx`) et Excel (`parametres_mis_a_jour.xlsx`) concernés sont **fermés**.
>
> En effet, la plupart des systèmes d’exploitation bloquent l’accès en écriture — et parfois même en lecture — à ces fichiers lorsqu’ils sont ouverts dans une application comme Microsoft Word ou Excel.

#### Lancer le script principal :

```bash
cd src
python main.py
```

### ✅ Résultat attendu :
- Le fichier Excel est mis à jour avec une nouvelle feuille versionnée
- Les cellules sont modifiées ou supprimées selon les couleurs détectées dans le Word
- Un message de confirmation s'affiche à la fin du traitement  

> [!TIP]
> Vous pouvez vous amuser à modifier les cellules surlignées dans le fichier Word d’entrée :
>
> - Changez la couleur d’une cellule en 🔴 rouge pour l’ajouter ou la modifier
> - Passez-la en 🌸 rose pour la supprimer
>
> Ensuite, relancez simplement le script `main.py` pour visualiser les effets dans le fichier Excel de sortie. Chaque exécution crée une nouvelle version dans une feuille dédiée, ce qui vous permet de suivre l’évolution des modifications pas à pas.
